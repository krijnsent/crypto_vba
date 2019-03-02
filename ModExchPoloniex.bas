Attribute VB_Name = "ModExchPoloniex"
Sub TestPoloniex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Poloniex will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_poloniex = "the key to use everywhere" etc )
apiKey = apikey_poloniex
secretKey = secretkey_poloniex

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchPoloniex"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestPoloniexPublic")

'Error, unknown command
TestResult = PublicPoloniex("returnUnknownCommand")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Invalid command."

'Error, missing parameters
TestResult = PublicPoloniex("returnOrderBook")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Please specify a currency pair."

'Testing error catching and replies
TestResult = PublicPoloniex("returnTicker")
'{"BTC_BCN":{"id":7,"last":"0.00000120","lowestAsk":"0.00000120","highestBid":"0.00000119","percentChange":"1.00000000","baseVolume":"21570.44763887","quoteVolume":"21082615430.89178085", etc...
Test.IsOk InStr(TestResult, "lowestAsk") > 0
Test.IsOk InStr(TestResult, "BTC_ETH"":") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), ""
Test.IsEqual JsonResult("BTC_ETH")("id"), 148
Test.IsOk JsonResult("BTC_ETH")("highestBid") > 0

TestResult = PublicPoloniex("returnOrderBook", "&currencyPair=BTC_ETH&depth=10")
'{"asks":[["0.05099419",0.14951192],["0.05099420",2.99201375],["0.05100000",28.07798797],["0.05101333",3.12600617],["0.05104000",13.17136949],["0.05104999",0.005],["0.05106858",0.2202525],["0.05107467",0.14672042],["0.05107609",0.44092991],["0.05108509",0.22025319]],"bids": etc...Test.IsEqual Left(TestResult, 40), "{""success"":true,""message"":"""",""result"":[{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), ""
'Test.IsEqual JsonResult("result")(1).Count, 7
'Test.IsOk JsonResult("result")(1)("Id") > 151

'Unix time period:
Set Test = Suite.Test("TestPoloniexPrivate")
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

TestResult = PrivatePoloniex("returnBalances", apiKey, secretKey)
'Debug.Print PrivatePoloniex("returnBalances", apiKey, secretkey)
''{"1CR":"0.00000000","ABY":"0.00000000","AC":"0.00000000","ACH":"0.00000000","ADN":"0.00000000","AEON":"0.00000000" etc...

'Debug.Print PrivatePoloniex("returnTradeHistory", apiKey, secretkey, "&currencyPair=all&start=" & t1 & "&end=" & t2)
''{"BTC_ETH":[{"globalTradeID":108848981,"tradeID":"22880801","date":"2017-04-19 23:26:55","rate":"0.03900000","amount":"65.35644222","total":"2.54890124", etc...

End Sub

Function PublicPoloniex(Method As String, Optional MethodOptions As String) As String

'https://poloniex.com/support/api/
Dim Url As String
PublicApiSite = "https://poloniex.com"
urlPath = "/public?command=" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicPoloniex = WebRequestURL(Url, "GET")

End Function
Function PrivatePoloniex(Method As String, apiKey As String, secretKey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
Dim postdata As String

'https://poloniex.com/support/api/

'Poloniex nonce
NonceUnique = CreateNonce(16)

Url = "https://poloniex.com/tradingApi"
postdata = "command=" & Method & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", postdata, secretKey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'If you get VBA: An error occurred in the secure channel support
'Check out: https://github.com/krijnsent/crypto_vba/issues/25 -> try the extra option below
'objHTTP.Option(4) = 13056
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Key", apiKey
objHTTP.setRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivatePoloniex = objHTTP.responseText
Set objHTTP = Nothing

End Function
