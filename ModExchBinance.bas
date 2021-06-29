Attribute VB_Name = "ModExchBinance"
Sub TestBinance()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://github.com/binance/binance-spot-api-docs/blob/master/rest-api.md
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_binance2
secretKey = secretkey_binance2

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBinance"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBinancePublic")

'Error, unknown command
TestResult = PublicBinance("AnUnknownCommand", "GET")
Test.IsOk InStr(TestResult, "error") > 0, "test UnknownCommand 1a failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test UnknownCommand 1b failed, result: ${1}"

'Error, command without parameters
TestResult = PublicBinance("depth", "GET")
Test.IsOk InStr(TestResult, "error") > 0, "test MissingParams 1a failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400, "test MissingParams 1b failed, result: ${1}"

'OK request
TestResult = PublicBinance("time", "GET")
'{"serverTime":1513605418615}
Test.IsOk InStr(TestResult, "serverTime") > 0, "test Time 1a failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("serverTime") > 1510000000000#, "test Time 1b failed, result: ${1}"

'Put parameters/options in a dictionary
Dim Params As New Dictionary
Params.Add "symbol", "ETHBTC"
TestResult = PublicBinance("ticker/24hr", "GET", Params)
'{"symbol":"ETHBTC","priceChange":"-0.00022700","priceChangePercent":"-0.633","weightedAvgPrice":"0.03538261","prevClosePrice":"0.03586800","lastPrice":"0.03564100","lastQty":"0.14000000","bidPrice":"0.03564100","bidQty":"0.22300000","askPrice":"0.03564800","askQty":"0.43200000","openPrice":"0.03586800","highPrice":"0.03600300","lowPrice":"0.03410000","volume":"380396.97600000","quoteVolume":"13459.43958266","openTime":1551288592637,"closeTime":1551374992637,"firstId":109505628,"lastId":109773015,"count":267388}
Test.IsOk InStr(TestResult, "priceChange") > 0, "test Ticker 1a failed, result: ${1}"
Test.IsOk InStr(TestResult, "closeTime") > 0, "test Ticker 1b failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("symbol"), "ETHBTC", "test Ticker 1c failed, result: ${1}"
Test.IsOk JsonResult("lastPrice") > 0, "test Ticker 1d failed, result: ${1}"

TestResult = GetBinanceTime()
'e.g. 1516565004894
Test.IsOk TestResult > 1510000000000#, "test GetTime failed, result: ${1}"

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Set Test = Suite.Test("TestBinancePrivate GET")
'Binance always requires a timestamp parameter, first test without
TestResult = PrivateBinance("api/v3/account", "GET", Cred)
'{"code":-1102,"msg":"Mandatory parameter 'timestamp' was not sent, was empty/null, or malformed."}
Test.IsOk InStr(TestResult, "code") > 0, "test Private GET 1a failed, result: ${1}"
Test.IsOk InStr(TestResult, "Mandatory parameter") > 0, "test Private GET 1b failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("response_txt")("code"), -1102, "test Private GET 1c failed, result: ${1}"

'Add timestamp to the parameters and try again
Dim Params2 As New Dictionary
Params2.Add "timestamp", GetBinanceTime()
TestResult = PrivateBinance("api/v3/account", "GET", Cred, Params2)
'{"makerCommission":10,"takerCommission":10,"buyerCommission":0,"sellerCommission":0,"canTrade":true,"canWithdraw":true,"canDeposit":true,"updateTime":1512476238993,"balances":[{"asset":"BTC","free":"0.00000000","locked":"0.00000000"},{"asset":"LTC","free":"0.00000000","locked":"0.00000000"},{"asset":"ETH","free":"0.00000000","locked":"0.00000000"},{"asset":"NEO","free":"0.00000000","locked":"0.00000000"},{"asset":"BNB","free":"0.00000000","locked":"0.00000000"},{"asset":"QTUM","free":"0.00000000","locked":"0.00000000"},{"asset":"EOS","free":"0.00000000","locked":"0.00000000"},{"asset":"SNT","free":"0.00000000","locked":"0.00000000"},{"asset":"BNT","free":"0.00000000","locked":"0.00000000"},{"asset":"GAS","free":"0.00000000","locked":"0.00000000"},{"asset":"BCC","free":"0.00000000","locked":"0.00000000"},{"asset":"USDT","free":"0.00000000","locked":"0.00000000"},{"asset":"HSR","free":"0.00000000","locked":"0.00000000"},{"asset":"OAX","free":"0.00000000","locked":"0.00000000"},{...
Test.IsOk InStr(TestResult, "takerCommission") > 0, "test Private GET 1d failed, result: ${1}"
Test.IsOk InStr(TestResult, "locked") > 0, "test Private GET 1e failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("takerCommission") > 0, "test Private GET 1f failed, result: ${1}"
Test.IsOk JsonResult("balances").Count > 10, "test Private GET 1g failed, result: ${1}"

Set Test = Suite.Test("TestBinancePrivate POST/DELETE")
'Test a test order
Dim Params3 As New Dictionary
Params3.Add "symbol", "LTCBTC"
Params3.Add "side", "BUY"
Params3.Add "type", "LIMIT"
Params3.Add "price", 0.01
Params3.Add "quantity", 1
Params3.Add "timeInForce", "GTC"
Params3.Add "timestamp", GetBinanceTime()
TestResult = PrivateBinance("api/v3/order/test", "POST", Cred, Params3)
Test.IsEqual TestResult, "{}", "test Private POST order 1a failed, result: ${1}"

'Delete a non-existing order
Dim Params4 As New Dictionary
Params4.Add "symbol", "LTCBTC"
Params4.Add "orderId", 987654
Params4.Add "timestamp", GetBinanceTime()
TestResult = PrivateBinance("api/v3/order", "DELETE", Cred, Params4)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"code":-2011,"msg":"Unknown order sent."}}
Test.IsOk InStr(TestResult, "code") > 0, "test Private DELETE order 1a failed, result: ${1}"
Test.IsOk InStr(TestResult, "Unknown order") > 0, "test Private DELETE order 1b failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("response_txt")("code"), -2011, "test Private DELETE order 1c failed, result: ${1}"


'Use the Wallet end point
Dim Params5 As New Dictionary
Params5.Add "timestamp", GetBinanceTime()
TestResult = PrivateBinance("sapi/v1/system/status", "GET", Cred, Params5)
'{"status":0,"msg":"normal"}
Test.IsOk InStr(TestResult, "msg") > 0, "test Private System Status 1a failed, result: ${1}"
Test.IsOk InStr(TestResult, "status") > 0, "test Private System Status 1b failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("msg"), "normal", "test Private System Status 1c failed, result: ${1}"

Dim Params6 As New Dictionary
Params6.Add "timestamp", GetBinanceTime()
TestResult = PrivateBinance("sapi/v1/capital/withdraw/history", "GET", Cred, Params6)
'e.g. [] (none) or
'e.g. [{"address":"0x94df8b352de7f46f64b01d3666bf6e936e44ce60","amount":"8.91000000","applyTime":"2019-10-1211:12:02","coin":"USDT","id":"b6ae22b3aa844210a7041aee7589627c","withdrawOrderId":"WITHDRAWtest123",//willnotbereturnedifthere'snowithdrawOrderIdforthiswithdraw."network":"ETH","transferType":0,//1forinternaltransfer,0forexternaltransfer"status":6,"transactionFee":"0.004","txId":"0xb5ef8c13b968a406cc62a93a8bd80f9e9a906ef1b3fcf20a2e48573c17659268"},{"address":"1FZdVHtiBqMrWdjPyRPULCUceZPJ2WLCsB","amount":"0.00150000","applyTime":"2019-09-2412:43:45","coin":"BTC","id":"156ec387f49b41df8724fa744fa82719","network":"BTC","status":6,"transactionFee":"0.004","transferType":0,//1forinternaltransfer,0forexternaltransfer"txId":"60fd9007ebfddc753455f95fafa808c4302c836e4d1eebc5a132c36c1d8ac354"}]
If TestResult = "[]" Then
    'Empty result, OK
    Test.IsEqual TestResult, "[]", "test Private Withdraw History 1a failed, result: ${1}"
Else
    Test.IsOk InStr(TestResult, "address") > 0, "test Private Withdraw History 1b failed, result: ${1}"
    Test.IsOk InStr(TestResult, "coin") > 0, "test Private Withdraw History 1c failed, result: ${1}"
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult(1)("amount") * 1 > 0, "test Private Withdraw History 1d failed, result: ${1}"
End If

End Sub

Function PublicBinance(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.binance.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/api/v1/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBinance = WebRequestURL(url, ReqType)

End Function
Function PrivateBinance(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim TimeCorrection As Long
Dim url As String

TradeApiSite = "https://api.binance.com/"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "&" & MethodParams

APIsign = ComputeHash_C("SHA256", MethodParams, Credentials("secretKey"), "STRHEX")
url = TradeApiSite & Method & "?" & MethodParams & "&signature=" & APIsign

Dim UrlHeaders As New Dictionary
UrlHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
UrlHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
UrlHeaders.Add "X-MBX-APIKEY", Credentials("apiKey")
PrivateBinance = WebRequestURL(url, ReqType, UrlHeaders)

End Function

Function GetBinanceTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicBinance time
JsonResponse = PublicBinance("time", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetBinanceTime = Json("serverTime")

Set Json = Nothing

End Function

