Attribute VB_Name = "ModExchPoloniex"
Sub TestPoloniex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'https://docs.poloniex.com/#http-api
'Poloniex will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_poloniex = "the key to use everywhere" etc )
apiKey = apikey_poloniex
secretKey = secretkey_poloniex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

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
TestResult = PublicPoloniex("returnUnknownCommand", "GET")
'{"error":"Invalid command."}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Invalid command."

'Error, missing parameters
TestResult = PublicPoloniex("returnOrderBook", "GET")
'{"error":"Please specify a currency pair."}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), "Please specify a currency pair."

'Testing error catching and replies
TestResult = PublicPoloniex("returnTicker", "GET")
'{"BTC_BCN":{"id":7,"last":"0.00000120","lowestAsk":"0.00000120","highestBid":"0.00000119","percentChange":"1.00000000","baseVolume":"21570.44763887","quoteVolume":"21082615430.89178085", etc...
Test.IsOk InStr(TestResult, "lowestAsk") > 0
Test.IsOk InStr(TestResult, "BTC_ETH"":") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error"), ""
Test.IsEqual JsonResult("BTC_ETH")("id"), 148
Test.IsOk JsonResult("BTC_ETH")("highestBid") > 0

'Put the parameters in a dictionary
Dim Params As New Dictionary
Params.Add "currencyPair", "BTC_ETH"
Params.Add "depth", 10
TestResult = PublicPoloniex("returnOrderBook", "GET", Params)
'{"asks":[["0.03530499",1.18647302],["0.03530500",110.78279995],["0.03531880",0.70796807],["0.03534095",2.12187844],["0.03534099",0.11553593],["0.03534767",29.95566069],["0.03534768",3.99999999],["0.03535000",0.99900001],["0.03535497",14.16571992],["0.03535498",0.6221801]],"bids":[["0.03528822",0.0031],["0.03528813",0.06749181],["0.03528730",0.0674917],["0.03528711",0.0674917],["0.03528638",0.0673596],["0.03528531",0.01],["0.03528303",0.01417112],["0.03527231",16.12158867],["0.03527000",110.5868],["0.03526147",33.74922032]],"isFrozen":"0","seq":644421713}
Test.IsOk InStr(TestResult, "],[") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("asks").Count, 10
Test.IsEqual JsonResult("bids").Count, 10
Test.IsOk JsonResult("asks")(1)(2) > 0

'Unix time period:
Set Test = Suite.Test("TestPoloniexPrivate")
t1 = DateToUnixTime("1/1/2016")
t2 = DateToUnixTime("1/1/2019")

TestResult = PrivatePoloniex("returnBalances", "POST", Cred)
'{"1CR":"0.00000000","ABY":"0.00000000","AC":"0.00000000","ACH":"0.00000000","ADN":"0.00000000","AEON":"0.00000000" etc...
Test.IsOk InStr(TestResult, "BTC") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count >= 10
Test.IsOk JsonResult("ETH") >= 0

'Put the parameters in a dictionary
Dim Params2 As New Dictionary
Params2.Add "currencyPair", "all"
Params2.Add "start", t1
Params2.Add "end", t2

TestResult = PrivatePoloniex("returnTradeHistory", "POST", Cred, Params2)
If InStr(TestResult, "globalTradeID") > 0 Then
    'has some results
    'e.g.: {"BTC_ETH":[{"globalTradeID":108848981,"tradeID":"22880801","date":"2017-04-19 23:26:55","rate":"0.03900000","amount":"65.35644222","total":"2.54890124", etc...
    Test.IsOk InStr(TestResult, "amount") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    For Each k In JsonResult.Keys()
        Test.IsOk JsonResult(k).Count >= 1
    Next k
Else
    'no results
    'Empty: []
    Test.IsEqual TestResult, "[]"
End If

'Put the parameters in a dictionary
Dim Params3 As New Dictionary
Params3.Add "currencyPair", "BTC_ETH"
Params3.Add "rate", 0.001
Params3.Add "amount", 3
Params3.Add "fillOrKill", 1

TestResult = PrivatePoloniex("buy", "POST", Cred, Params3)
'{"error":"This API key does not have permission to trade."}
'{orderNumber: '514845991795',resultingTrades:[{amount: '3.0',Date: '2018-10-25 23:03:21',rate:'0.0002',total:'0.0006',tradeID:'251834',type:'buy'}]}
If InStr(TestResult, "error") > 0 Then
    Test.IsOk InStr(TestResult, "permission") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsEqual JsonResult("error"), "This API key does not have permission to trade."
Else
    Test.IsOk InStr(TestResult, "resultingTrades") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk JsonResult("orderNumber") >= 0
End If


End Sub

Function PublicPoloniex(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://poloniex.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "&" & MethodParams
urlPath = "/public?command=" & Method & MethodParams
Url = PublicApiSite & urlPath

PublicPoloniex = WebRequestURL(Url, ReqType)

End Function
Function PrivatePoloniex(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim PayloadDict As Dictionary

'Poloniex nonce
NonceUnique = CreateNonce(16)

Url = "https://poloniex.com/tradingApi"

Set PayloadDict = New Dictionary
PayloadDict("command") = Method
If Not ParamDict Is Nothing Then
    For Each Key In ParamDict.Keys
        PayloadDict(Key) = ParamDict(Key)
    Next Key
End If
PayloadDict("&nonce") = NonceUnique

postdata = DictToString(PayloadDict, "URLENC")
APIsign = ComputeHash_C("SHA512", postdata, Credentials("secretKey"), "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'If you get VBA: An error occurred in the secure channel support
'Check out: https://github.com/krijnsent/crypto_vba/issues/25 -> try the extra option below
'objHTTP.Option(4) = 13056
objHTTP.Open ReqType, Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Key", Credentials("apiKey")
objHTTP.setRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivatePoloniex = objHTTP.responseText
Set objHTTP = Nothing

End Function
