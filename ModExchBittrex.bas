Attribute VB_Name = "ModExchBittrex"
Sub TestBittrex()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://bittrex.com/home/api
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

'Error, unknown command
TestResult = PublicBittrex("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, parameter missing
TestResult = PublicBittrex("getticker", "GET")
'{"success":false,"message":"MARKET_NOT_PROVIDED","result":null}
Test.IsOk InStr(TestResult, "message") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), False
Test.IsEqual JsonResult("message"), "MARKET_NOT_PROVIDED"

'Request without parameters
TestResult = PublicBittrex("getmarkets", "GET")
'{"success":true,"message":"","result":[{"MarketCurrency":"LTC","BaseCurrency":"BTC","MarketCurrencyLong":"Litecoin","BaseCurrencyLong":"Bitcoin","MinTradeSize":0.02103262,"MarketName":"BTC-LTC","IsActive":true, etc
Test.IsOk InStr(TestResult, "MarketCurrency") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("result")(1)("MarketCurrency"), "LTC"
Test.IsEqual JsonResult("result")(1)("BaseCurrency"), "BTC"

'Put parameters/options in a dictionary for a summary of one coin
Dim Params As New Dictionary
Params.Add "market", "BTC-DOGE"
TestResult = PublicBittrex("getmarketsummary", "GET", Params)
'{"success":true,"message":"","result":[{"MarketName":"BTC-DOGE","High":0.00000052,"Low":0.00000050,"Volume":53181123.38404553,"Last":0.00000051,"BaseVolume":27.40729560,"TimeStamp":"2019-03-02T17:17:08.867","Bid":0.00000051,"Ask":0.00000052,"OpenBuyOrders":1209,"OpenSellOrders":3081,"PrevDay":0.00000050,"Created":"2014-02-13T00:00:00"}]}
Test.IsOk InStr(TestResult, "MarketName") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("result")(1)("MarketName"), "BTC-DOGE"
Test.IsOk JsonResult("result")(1)("Volume") > 0

Set Test = Suite.Test("TestBittrexPrivate")

TestResult = PrivateBittrex("account/getbalances", "GET", Cred)
'{"success":true,"message":"","result":[{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"},{"Currency":"XMR","Balance":0.00000000,"Available":0.00000000,"Pending":0.00000000,"CryptoAddress":etc...
Test.IsOk InStr(TestResult, "result") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsOk Len(JsonResult("result")(1)("Currency")) >= 3
Test.IsOk JsonResult("result")(1)("Balance") > 0


Dim Params2 As New Dictionary
Params2.Add "currency", "ETH"
TestResult = PrivateBittrex("account/getdeposithistory", "GET", Cred, Params2)
'{"success":true,"message":"","result":[{"Id":44147323,"Amount":0.09430706,"Currency":"ETH","Confirmations":44,"LastUpdated":"2017-10-05T03:37:42.16","TxId":"0x870efe7ca6bca4","CryptoAddress":"0x6f61c"}]}
'or if no deposit was made: {"success":true,"message":"","result":[]}
Test.IsOk InStr(TestResult, "result") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsOk JsonResult("result").Count >= 0


End Sub

Function PublicBittrex(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://bittrex.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/api/v1.1/public/" & Method & MethodParams
Url = PublicApiSite & urlPath

PublicBittrex = WebRequestURL(Url, ReqType)

End Function
Function PrivateBittrex(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim Url As String

'Get a 10-digit Nonce
NonceUnique = CreateNonce(10)
TradeApiSite = "https://bittrex.com/api/v1.1/"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "&" & MethodParams

postdata = Method & "?apikey=" & Credentials("apiKey") & MethodParams & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", TradeApiSite & postdata, Credentials("secretKey"), "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "apisign", APIsign

Url = TradeApiSite & postdata
PrivateBittrex = WebRequestURL(Url, ReqType, headerDict)

End Function

