Attribute VB_Name = "ModExchBittrex"
Sub TestBittrex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretkey As String

apiKey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_bittrex
secretkey = secretkey_bittrex

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBittrex"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBittrexPublic")

'Testing error catching and replies
TestResult = PublicBittrex("getmarketsummary", "?market=btc-DOGE")
'{"success":true,"message":"","result":[{"MarketName":"BTC-LTC","High":0.01250680,"Low":0.01132497,"Volume":222923.75389408,"Last":0.01218025,"BaseVolume":2639.03223291,"TimeStamp":"2017-06-15T20:49:50.27","Bid":0.01218026,"Ask":0.01224870,"OpenBuyOrders":1439,"OpenSellOrders":2785,"PrevDay":0.01137500,"Created":"2014-02-13T00:00:00"}]}
Test.IsEqual Left(TestResult, 40), "{""success"":true,""message"":"""",""result"":[{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("message"), ""
Test.IsEqual JsonResult("result")(1)("MarketName"), "BTC-DOGE"
Test.IsOk JsonResult("result")(1)("TimeStamp") > 151


TestResult = PublicBittrex("getmarkethistory", "?market=BTC-DOGE")
'{"success":true,"message":"","result":[{"Id":6313536,"TimeStamp":"2017-06-15T20:49:05.46","Quantity":84553.23767320,etc.
Test.IsEqual Left(TestResult, 40), "{""success"":true,""message"":"""",""result"":[{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("message"), ""
Test.IsEqual JsonResult("result")(1).Count, 7
Test.IsOk JsonResult("result")(1)("Id") > 151


'Unix time period:
Set Test = Suite.Test("TestBittrexPrivate")
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

TestResult = PrivateBittrex("account/getbalances", apiKey, secretkey)
'{"success":true,"message":"","result":[{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"},{"Currency":"XMR","Balance":0.00000000,"Available":0.00000000,"Pending":0.00000000,"CryptoAddress":etc...
Test.IsEqual Left(TestResult, 40), "{""success"":true,""message"":"""",""result"":[{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("message"), ""
Test.IsEqual JsonResult("result")(1).Count, 5
Test.IsOk Len(JsonResult("result")(1)("Currency")) > 0
Test.IsOk JsonResult("result")(1)("Balance") >= 0


TestResult = PrivateBittrex("account/getbalance", apiKey, secretkey, "&currency=ETH")
'{"success":true,"message":"","result":{"Currency":"ETH","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"}}
Test.IsEqual Left(TestResult, 39), "{""success"":true,""message"":"""",""result"":{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("message"), ""
Test.IsOk JsonResult("result").Count > 0
Test.IsOk Len(JsonResult("result")("Currency")) > 0


End Sub

Function PublicBittrex(Method As String, Optional MethodOptions As String) As String

'https://bittrex.com/home/api
Dim Url As String
PublicApiSite = "https://bittrex.com"
urlPath = "/api/v1.1/public/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicBittrex = WebRequestURL(Url, "GET")

End Function
Function PrivateBittrex(Method As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
Dim postdata As String
Dim Url As String
'https://bittrex.com/home/api

'Get a 10-digit Nonce
NonceUnique = CreateNonce(10)
TradeApiSite = "https://bittrex.com/api/v1.1/"

postdata = Method & "?apikey=" & apiKey & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", TradeApiSite & postdata, secretkey, "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "apisign", APIsign

Url = TradeApiSite & postdata
PrivateBittrex = WebRequestURL(Url, "POST", headerDict)

End Function

