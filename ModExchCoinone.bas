Attribute VB_Name = "ModExchCoinone"
Sub TestCoinone()

'Source: https://github.com/krijnsent/crypto_vba
'https://doc.coinone.co.kr/#section/V2-version
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_coinone
secretKey = secretkey_coinone

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchCoinone"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestCoinonePublic")

'Error, unknown command
TestResult = PublicCoinone("AnUnknownCommand", "GET")
'{"error_nr":200,"error_txt":"NO JSON BUT HTML RETURNED","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_txt"), "HTTP-Not Found"
Test.IsEqual JsonResult("error_nr"), 404

'OK request
TestResult = PublicCoinone("ticker", "GET")
'e.g. {"currency":"btc","volume":"633.1048","last":"4684000.0","yesterday_last":"4636000.0","timestamp":"1554107620","yesterday_low":"4592000.0","errorCode":"0","yesterday_volume":"395.8966","high":"4720000.0","result":"success","yesterday_first":"4615000.0","first":"4636000.0","yesterday_high":"4651000.0","low":"4630000.0"}
Test.IsOk InStr(TestResult, "yesterday_last") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Val(JsonResult("last")) > 0
Test.IsEqual JsonResult("currency"), "btc"
Test.IsOk Val(JsonResult("timestamp")) > 1500000000#

'Put parameters/options in a dictionary
'If no parameters are provided, the defaults are used
'If WRONG PARAMETERS are provided, the defaults will be used: the API fails "silently" and gives no error but default BTC data
Dim Params As New Dictionary
Params.Add "currency", "eth"
Params.Add "period", "hour"
TestResult = PublicCoinone("trades", "GET", Params)
'e.g. {"errorCode":"0","timestamp":"1554107995","completeOrders":[{"is_ask":"0","timestamp":"1554107949","price":"161600.0","id":"395377","qty":"1.6044"},
Test.IsOk InStr(TestResult, "completeOrders") > 0
Test.IsOk InStr(TestResult, "timestamp") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Val(JsonResult("timestamp")) > 1500000000#
Test.IsEqual JsonResult("errorCode"), "0"
Test.IsEqual JsonResult("completeOrders").Count, 200
Test.IsOk Val(JsonResult("completeOrders")(1)("id")) > 0
Test.IsOk Val(JsonResult("completeOrders")(1)("qty")) > 0


Set Test = Suite.Test("TestCoinonePrivate")

TestResult = PrivateCoinone2("account/balance", "POST", Cred)
'{"btt": {"avail": "0.0", "balance": "0.0"}, "edna": {"avail": "0.0", "balance": "0.0"},  etc.
Test.IsOk InStr(TestResult, "avail") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("result"), "success"
Test.IsEqual JsonResult("errorCode"), "0"
Test.IsOk JsonResult("eos")("avail") >= 0
Test.IsOk JsonResult("btc")("balance") >= 0

Dim Params2 As New Dictionary
Params2.Add "price", 100
Params2.Add "qty", 3
Params2.Add "currency", "EOS"
TestResult = PrivateCoinone2("order/limit_buy", "POST", Cred, Params2)
'{"errorCode":"103","errorMsg":"Lack of Balance","result":"error"}
'{"errorCode":"113","errorMsg":"Quantity is too low","result":"error"}
'{"result": "success","errorCode": "0","orderId": "8a82c561-40b4-4cb3-9bc0-9ac9ffc1d63b"}
Test.IsOk InStr(TestResult, "errorCode") > 0
Test.IsOk InStr(TestResult, "result") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
If Val(JsonResult("errorCode")) = 0 Then
    'No error
    Test.IsEqual JsonResult("result"), "success"
    Test.IsEqual JsonResult("errorCode"), "0"
    Test.IsOk Len(JsonResult("orderId")) > 10
Else
    'Error
    Test.IsEqual JsonResult("result"), "error"
    Test.IsOk Len(JsonResult("errorMsg")) > 0
End If

Dim Params3 As New Dictionary
Params3.Add "currency", "ETH"
TestResult = PrivateCoinone2("order/complete_orders", "POST", Cred, Params3)
'{"errorCode": "0", "completeOrders": [], "result": "success"}
'{"result": "success","errorCode": "0","completeOrders": [{"timestamp": "1416561032","price": "419000.0","type": "bid","qty": "0.001","feeRate": "-0.0015","fee": "-0.0000015","orderId": "E84A1AC2-8088-4FA0-B093-A3BCDB9B3C85"}]}
Test.IsOk InStr(TestResult, "completeOrders") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("result"), "success"
Test.IsEqual JsonResult("errorCode"), "0"
Test.IsOk JsonResult("completeOrders").Count >= 0

End Sub

Function PublicCoinone(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://api.coinone.co.kr/"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = Method & MethodParams
Url = PublicApiSite & urlPath

PublicCoinone = WebRequestURL(Url, ReqType)

End Function
Function PrivateCoinone2(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim postdataUrl As String
Dim postdataJSON As String
Dim Url As String

'Get a 14-digit Nonce
NonceUnique = CreateNonce(14)
TradeApiSite = "https://api.coinone.co.kr/v2/"

Url = TradeApiSite & Method

Dim PostDict As New Dictionary
PostDict.Add "access_token", Credentials("apiKey")
PostDict.Add "nonce", NonceUnique
If Not ParamDict Is Nothing Then
    For Each Key In ParamDict.Keys
        PostDict(Key) = ParamDict(Key)
    Next Key
End If

postdataUrl = DictToString(PostDict, "URLENC")
postdataJSON = DictToString(PostDict, "JSON")
postdata64 = Base64Encode(postdataJSON)

APIsign = ComputeHash_C("SHA512", postdata64, Credentials("secretKey"), "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/json"
headerDict.Add "X-COINONE-PAYLOAD", postdata64
headerDict.Add "X-COINONE-SIGNATURE", APIsign

Url = TradeApiSite & Method
PrivateCoinone2 = WebRequestURL(Url, ReqType, headerDict, postdataUrl)

End Function
