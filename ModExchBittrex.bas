Attribute VB_Name = "ModExchBittrex"
Sub TestBittrex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apikey = apikey_bittrex
secretkey = secretkey_bittrex

Debug.Print PublicBittrex("getmarketsummary", "?market=btc-DOGE")
'{"success":true,"message":"","result":[{"MarketName":"BTC-LTC","High":0.01250680,"Low":0.01132497,"Volume":222923.75389408,"Last":0.01218025,"BaseVolume":2639.03223291,"TimeStamp":"2017-06-15T20:49:50.27","Bid":0.01218026,"Ask":0.01224870,"OpenBuyOrders":1439,"OpenSellOrders":2785,"PrevDay":0.01137500,"Created":"2014-02-13T00:00:00"}]}
Debug.Print PublicBittrex("getmarkethistory", "?market=BTC-DOGE")
'{"success":true,"message":"","result":[{"Id":6313536,"TimeStamp":"2017-06-15T20:49:05.46","Quantity":84553.23767320,etc.

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateBittrex("account/getbalances", apikey, secretkey)
'{"success":true,"message":"","result":[{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"},{"Currency":"XMR","Balance":0.00000000,"Available":0.00000000,"Pending":0.00000000,"CryptoAddress":etc...
Debug.Print PrivateBittrex("account/getbalance", apikey, secretkey, "&currency=ETH")
'{"success":true,"message":"","result":{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"}}

End Sub

Function PublicBittrex(Method As String, Optional MethodOptions As String) As String

'https://bittrex.com/home/api
Dim Url As String
PublicApiSite = "https://bittrex.com"
urlPath = "/api/v1.1/public/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicBittrex = GetDataFromURL(Url, "GET")

End Function
Function PrivateBittrex(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'https://bittrex.com/home/api

'Get a 10-digit Nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)
TradeApiSite = "https://bittrex.com/api/v1.1/"

postdata = Method & "?apikey=" & apikey & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", TradeApiSite & postdata, secretkey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'Debug.Print "POST: " & TradeApiSite & postdata
objHTTP.Open "POST", TradeApiSite & postdata, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "apisign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateBittrex = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

