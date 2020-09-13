Attribute VB_Name = "ModExchBittrex"
Sub TestBittrex()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://bittrex.com/home/api
'v3 - https://bittrex.github.io/api/v3
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_bittrex
secretKey = secretkey_bittrex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBittrex"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBittrexPublic")

'Error, unknown/wrong command
TestResult = PublicBittrex("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"code":"NOT_FOUND"}}
Test.IsOk InStr(TestResult, "error") > 0, "test error 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test error 2 failed, result: ${1}"

'Request without parameters
TestResult = PublicBittrex("markets", "GET")
'[{"symbol":"4ART-BTC","baseCurrencySymbol":"4ART","quoteCurrencySymbol":"BTC","minTradeSize":"10.00000000","precision":8,"status":"ONLINE","createdAt":"2020-06-10T15:05:29.833Z","notice":"","prohibitedIn":["US"],"associatedTermsOfService":[]},{"symbol":"4ART-USDT","baseCurrencySymbol":"4ART","quoteCurrencySymbol":"USDT","minTradeSize":"10.00000000","precision":5,"status":"ONLINE","createdAt":"2020-06-10T15:05:40.98Z", etc.
Test.IsOk InStr(TestResult, "quoteCurrencySymbol") > 0, "test markets 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)("quoteCurrencySymbol"), "BTC", "test markets 2 failed, result: ${1}"
Test.IsOk JsonResult(1)("precision") > 0, "test markets 3 failed, result: ${1}"

'Put parameters/options in a dictionary for a summary of one coin, wrong input
Dim Params As New Dictionary
Params.Add "market", "BTC-DOGE"
TestResult = PublicBittrex("markets", "GET", Params)
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"code":"MARKET_NAME_REVERSED","detail":"The provided market symbol appears to be reversed. Please retry with the market symbol provided in data.NewMarketSymbol.","data":{"newMarketSymbol":"DOGE-BTC"}}}
Test.IsOk InStr(TestResult, "error") > 0, "test error2 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test error2 2 failed, result: ${1}"

'Parameter in a dictionary
Dim Params2 As New Dictionary
Params2.Add "market", "DOGE-BTC"
TestResult = PublicBittrex("markets", "GET", Params2)
'{"symbol":"DOGE-BTC","baseCurrencySymbol":"DOGE","quoteCurrencySymbol":"BTC","minTradeSize":"1000.00000000","precision":8,"status":"ONLINE","createdAt":"2014-02-13T00:00:00Z","prohibitedIn":[],"associatedTermsOfService":[]}
Test.IsOk InStr(TestResult, "baseCurrencySymbol") > 0, "test markets detail 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("quoteCurrencySymbol"), "BTC", "test markets detail 2 failed, result: ${1}"
Test.IsEqual JsonResult("precision"), 8, "test markets detail 3 failed, result: ${1}"
Test.IsEqual JsonResult("baseCurrencySymbol"), "DOGE", "test markets detail 4 failed, result: ${1}"

'Parameters in a dictionary get '/markets/{marketSymbol}/candles/{candleInterval}/recent
Dim Params3 As New Dictionary
Params3.Add "market", "ETH-BTC"
Params3.Add "type1", "candles"
Params3.Add "candleInterval", "HOUR_1"
Params3.Add "type2", "recent"
TestResult = PublicBittrex("markets", "GET", Params3)
'[{"startsAt":"2020-08-13T15:00:00Z","open":"0.03405607","high":"0.03412946","low":"0.03393712","close":"0.03411082","volume":"224.62409851","quoteVolume":"7.64110651"},{"startsAt":"2020-08-13T16:00:00Z","open":"0.03411095","high":"0.03418634","low":"0.03387446","close":"0.03402789","volume":"303.55027355","quoteVolume":"10.33201616"},{"startsAt":"2020-08-13T17:00:00Z","open":"0.03403607","high":"0.03407806","low":"0.03389236","close":"0.03403147","volume":"487.61617145","quoteVolume":"16.57089220"},{"startsAt":"2020-08-13T18:00:00Z","open":"0.03403252","high":"0.03413220","low":"0.03403252","close":"0.03410964","volume":"388.13757692","quoteVolume":"13.22881730"},{"startsAt":"2020-08-13T19:00:00Z","open":"0.03408765","high":"0.03425485","low":"0.03408765","close":"0.03422712","volume":"312.75229144","quoteVolume":"10.69620756"}, etc...
Test.IsOk InStr(TestResult, "startsAt") > 0, "test candles 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult(1)("open") > 0, "test candles 2 failed, result: ${1}"
Test.IsOk JsonResult(1)("high") > 0, "test candles 3 failed, result: ${1}"
Test.IsOk JsonResult(1)("low") > 0, "test candles 4 failed, result: ${1}"

'Get bittrex time from ping
Set Test = Suite.Test("TestBittrexTime")
TestResult = GetBittrexTime()
Test.IsOk TestResult > 0, "test time 2 failed, result: ${1}"

'Test private API
Set Test = Suite.Test("TestBittrexPrivate")
TestResult = PrivateBittrex("balances", "GET", Cred)
'[{"currencySymbol":"BCH","total":"0.00001733","available":"0.00001733","updatedAt":"2001-01-01T00:00:00Z"},{"currencySymbol":"BTC","total":"0.01500039","available":"0.01500039","updatedAt":"2001-01-01T00:00:00Z"},{"currencySymbol":"BTXCRD","total":"0.00000000","available":"0.00000000","updatedAt":"2019-10-23T04:16:31.1Z"},{"currencySymbol":"XLM","total":"0","available":"0","updatedAt":"2020-09-13T16:02:42.84307Z"}], etc...
Test.IsOk InStr(TestResult, "currencySymbol") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Len(JsonResult(1)("currencySymbol")) >= 3
Test.IsOk JsonResult(1)("Balance") >= 0

'Test private API
Dim Params4 As New Dictionary
Params4.Add "marketSymbol", "DOGE-BTC"
Set Test = Suite.Test("TestBittrexPrivate")
TestResult = PrivateBittrex("orders/open", "GET", Cred, Params4)
Debug.Print TestResult

'[{"currencySymbol":"BCH","total":"0.00001733","available":"0.00001733","updatedAt":"2001-01-01T00:00:00Z"},{"currencySymbol":"BTC","total":"0.01500039","available":"0.01500039","updatedAt":"2001-01-01T00:00:00Z"},{"currencySymbol":"BTXCRD","total":"0.00000000","available":"0.00000000","updatedAt":"2019-10-23T04:16:31.1Z"},{"currencySymbol":"XLM","total":"0","available":"0","updatedAt":"2020-09-13T16:02:42.84307Z"}], etc...
Test.IsOk InStr(TestResult, "currencySymbol") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Len(JsonResult(1)("currencySymbol")) >= 3
Test.IsOk JsonResult(1)("Balance") >= 0


Dim Params5 As New Dictionary
Params5.Add "marketSymbol", "DOGE-BTC"
Params5.Add "direction", "BUY"
Params5.Add "type", "LIMIT"
Params5.Add "quantity", 10
Params5.Add "timeInForce", "FILL_OR_KILL"
Params5.Add "quantity", 10

TestResult = PrivateBittrex("orders", "POST", Cred, Params5)
'{"success":false,"message":"INVALID_ORDER","result":null}
Test.IsOk InStr(TestResult, "result") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), False
Test.IsOk JsonResult("message") = "INVALID_ORDER"

End Sub

Function PublicBittrex(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.bittrex.com"

MethodParams = ""
If Not ParamDict Is Nothing Then
    For Each itm In ParamDict
        MethodParams = MethodParams & ParamDict(itm) & "/"
    Next itm
End If

urlPath = "/v3/" & Method & "/" & MethodParams
url = PublicApiSite & urlPath

PublicBittrex = WebRequestURL(url, ReqType)

End Function
Function PrivateBittrex(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim url As String
Dim Uri As String

'Get a 13-digit Nonce from the server time
NonceUnique = GetBittrexTime
TradeApiSite = "https://api.bittrex.com/v3/"

Uri = TradeApiSite & Method
PostContent = DictToString(ParamDict, "URLENC")
'Uri = Uri & "/?" & PostContent
contentHash = ComputeHash_C("SHA512", PostContent, "", "STRHEX")
preSign = NonceUnique & Uri & ReqType & contentHash

'MethodParams = DictToString(ParamDict, "URLENC")
'If MethodParams <> "" Then MethodParams = "&" & MethodParams

'postdata = Method & "?apikey=" & Credentials("apiKey") & MethodParams & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", preSign, Credentials("secretKey"), "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "Api-Key", Credentials("apiKey")
headerDict.Add "Api-Timestamp", NonceUnique
headerDict.Add "Api-Content-Hash", contentHash
headerDict.Add "Api-Signature", APIsign

url = TradeApiSite & postdata
PrivateBittrex = WebRequestURL(Uri, ReqType, headerDict)

End Function


Function GetBittrexTime() As Double

Dim JsonResponse As String
Dim Json As Object

'GetBittrexTime time from ping
JsonResponse = PublicBittrex("ping", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetBittrexTime = Json("serverTime")
NonceUnique = CreateNonce(13)

If GetBittrexTime = 0 Then
    TimeCorrection = -3600
    GetBittrexTime = DateDiff("s", "1/1/1970", Now)
    GetBittrexTime = Trim(Str((Val(GetBittrexTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set Json = Nothing

End Function

