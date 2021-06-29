Attribute VB_Name = "ModExchBitstamp"
Sub TestBitstamp()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://Bitstamp.com/home/api
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"
customerID = "your customer id here"

'Remove these 3 lines, unless you define 3 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_bitstamp
secretKey = secretkey_bitstamp
customerID = customer_id_bitstamp

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey
Cred.Add "customerID", customerID

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBitstamp"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBitstampPublic")

'Error, unknown command
TestResult = PublicBitstamp("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, parameter missing
TestResult = PublicBitstamp("v2/ticker_hour/", "GET")
'{"error_nr":404,"error_txt":"HTTP-NOT FOUND","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Request without parameters
TestResult = PublicBitstamp("ticker/", "GET")
'{"high": "3806.90000000", "last": "3707.22", "timestamp": "1551731354", "bid": "3707.14", "vwap": "3724.51", "volume": "6515.58124105", "low": "3670.00000000", "ask": "3707.22", "open": 3789.70}
Test.IsOk InStr(TestResult, "timestamp") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("timestamp") * 1 >= 1510000000
Test.IsOk JsonResult("bid") >= 0

'Put variables in
TestResult = PublicBitstamp("v2/ticker_hour/btceur/", "GET")
'{"high": "3282.58", "last": "3277.18", "timestamp": "1551731355", "bid": "3276.00", "vwap": "3276.05", "volume": "24.42762265", "low": "3270.77", "ask": "3276.08", "open": "3275.17"}
Test.IsOk InStr(TestResult, "timestamp") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("timestamp") * 1 >= 1510000000
Test.IsOk JsonResult("bid") >= 0

'Unix time period:
Set Test = Suite.Test("TestBitstampPrivate")

TestResult = PrivateBitstamp("balance/", "POST", Cred)
'{"xrp_available": "0.00000000", "eur_available": "0.00", "usd_reserved": "0.00", "eur_balance": "0.00", "btc_balance": "0.00000000", "usd_available": "0.00", "btc_reserved": "0.00000000", "fee": "0.2500", "btc_available": "0.00000000", "eur_reserved": "0.00", "xrp_reserved": "0.00000000", "xrp_balance": "0.00000000", "usd_balance": "0.00"}
Test.IsOk InStr(TestResult, "eur_balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("eur_balance") >= 0
Test.IsOk JsonResult("usd_balance") >= 0

TestResult = PrivateBitstamp("v2/balance/", "POST", Cred)
'{"bch_available": "0.00000000", "bch_balance": "0.00000000", "bch_reserved": "0.00000000", "bchbtc_fee": "0.25", "bcheur_fee": "0.25", "bchusd_fee": "0.25", "btc_available": "0.00000000", "btc_balance": "0.00000000", "btc_reserved": "0.00000000", "btceur_fee": "0.25", "btcusd_fee": "0.25", "eth_available": "0.00000000", "eth_balance": "0.00000000", "eth_reserved": "0.00000000", "ethbtc_fee": "0.25", "etheur_fee": "0.25", "ethusd_fee": "0.25", "eur_available": "0.00", "eur_balance": "0.00", "eur_reserved": "0.00", "eurusd_fee": "0.25", "ltc_available": "0.00000000", "ltc_balance": "0.00000000", "ltc_reserved": "0.00000000", "ltcbtc_fee": "0.25", "ltceur_fee": "0.25", "ltcusd_fee": "0.25", "usd_available": "0.00", "usd_balance": "0.00", "usd_reserved": "0.00", "xrp_available": "0.00000000", "xrp_balance": "0.00000000", "xrp_reserved": "0.00000000", "xrpbtc_fee": "0.25", "xrpeur_fee": "0.25", "xrpusd_fee": "0.25"}
Test.IsOk InStr(TestResult, "btc_balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("bch_balance") >= 0
Test.IsOk JsonResult("eth_balance") >= 0

TestResult = PrivateBitstamp("order_status/", "POST", Cred)
'{"error": "Missing id POST param"}
Test.IsOk InStr(TestResult, "Missing") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Missing id POST param"

'Put the parameters in a dictionary
Dim Params As New Dictionary
Params.Add "id", 12345
TestResult = PrivateBitstamp("order_status/", "POST", Cred, Params)
'{"error": "Order not found"}
Test.IsOk InStr(TestResult, "found") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Order not found"


'Set a buy order, put the parameters in a dictionary
Dim Params2 As New Dictionary
Params2.Add "amount", 1
Params2.Add "price", 3
Params2.Add "ioc_order", True
TestResult = PrivateBitstamp("v2/buy/etheur/", "POST", Cred, Params2)
'{"status": "error", "reason": {"__all__": ["Minimum order size is 5.0 EUR."]}}
Test.IsOk InStr(TestResult, "Minimum order size") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "error"


End Sub

Function PublicBitstamp(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://www.bitstamp.net"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/api/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBitstamp = WebRequestURL(url, ReqType)

End Function
Function PrivateBitstamp(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim message As String
Dim PostMsg As String
Dim url As String
Dim PayloadDict As Dictionary

'Get a Nonce
NonceUnique = CreateNonce()
TradeApiSite = "https://www.bitstamp.net/api/"

message = NonceUnique & Credentials("customerID") & Credentials("apiKey")
APIsign = UCase(ComputeHash_C("SHA256", message, Credentials("secretKey"), "STRHEX"))

Set PayloadDict = New Dictionary
PayloadDict("key") = Credentials("apiKey")
PayloadDict("signature") = APIsign
PayloadDict("nonce") = NonceUnique
If Not ParamDict Is Nothing Then
    For Each key In ParamDict.Keys
        PayloadDict(key) = ParamDict(key)
    Next key
End If
PostMsg = DictToString(PayloadDict, "URLENC")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"

url = TradeApiSite & Method
PrivateBitstamp = WebRequestURL(url, ReqType, headerDict, PostMsg)

End Function


