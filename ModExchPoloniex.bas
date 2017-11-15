Attribute VB_Name = "ModExchPoloniex"
Sub TestPoloniex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Poloniex will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_poloniex = "the key to use everywhere" etc )
apikey = apikey_poloniex
secretkey = secretkey_poloniex

Debug.Print PublicPoloniex("returnTicker")
'{"BTC_BCN":{"id":7,"last":"0.00000120","lowestAsk":"0.00000120","highestBid":"0.00000119","percentChange":"1.00000000","baseVolume":"21570.44763887","quoteVolume":"21082615430.89178085", etc...
Debug.Print PublicPoloniex("returnOrderBook", "&currencyPair=BTC_ETH&depth=10")
'{"asks":[["0.05099419",0.14951192],["0.05099420",2.99201375],["0.05100000",28.07798797],["0.05101333",3.12600617],["0.05104000",13.17136949],["0.05104999",0.005],["0.05106858",0.2202525],["0.05107467",0.14672042],["0.05107609",0.44092991],["0.05108509",0.22025319]],"bids": etc...

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivatePoloniex("returnBalances", apikey, secretkey)
'{"1CR":"0.00000000","ABY":"0.00000000","AC":"0.00000000","ACH":"0.00000000","ADN":"0.00000000","AEON":"0.00000000" etc...
Debug.Print PrivatePoloniex("returnTradeHistory", apikey, secretkey, "&currencyPair=all&start=" & t1 & "&end=" & t2)
'{"BTC_ETH":[{"globalTradeID":108848981,"tradeID":"22880801","date":"2017-04-19 23:26:55","rate":"0.03900000","amount":"65.35644222","total":"2.54890124", etc...

End Sub

Function PublicPoloniex(Method As String, Optional MethodOptions As String) As String

'https://poloniex.com/support/api/

PublicApiSite = "https://poloniex.com"
urlPath = "/public?command=" & Method & MethodOptions
Url = PublicApiSite & urlPath

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", Url
objHTTP.Send
objHTTP.WaitForResponse
PublicPoloniex = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
Function PrivatePoloniex(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
Dim postdata As String

'https://poloniex.com/support/api/

'Poloniex nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)
NonceUnique = NonceUnique & Right(Timer * 100, 2) & "0000"

Url = "https://poloniex.com/tradingApi"
postdata = "command=" & Method & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", postdata, secretkey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.SetRequestHeader "Key", apikey
objHTTP.SetRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivatePoloniex = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
