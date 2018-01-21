Attribute VB_Name = "ModExchBinance"
Sub TestBinance()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apikey = apikey_binance
secretkey = secretkey_binance

Debug.Print PublicBinance("time", "")
'{"serverTime":1513605418615}
Debug.Print PublicBinance("ticker/24hr", "?symbol=ETHBTC")
'{"symbol":"ETHBTC","priceChange":"0.00231500","priceChangePercent":"6.345","weightedAvgPrice":"0.03788715","prevClosePrice":"0.03648400","lastPrice":"0.03880200","lastQty":"0.29800000","bidPrice":"0.03873300","bidQty":"10.00000000","askPrice":"0.03883100","askQty":"17.18000000","openPrice":"0.03648700","highPrice":"0.04000000","lowPrice":"0.03631200","volume":"274355.20000000","quoteVolume":"10394.53526717","openTime":1513522564335,"closeTime":1513608964335,"firstId":7427497,"lastId":7702400,"count":274904}
Debug.Print GetBinanceTime()
'e.g. 1516565004894

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateBinance("account", apikey, secretkey)
'{"makerCommission":10,"takerCommission":10,"buyerCommission":0,"sellerCommission":0,"canTra etc...
Debug.Print PrivateBinance("order/test", apikey, secretkey, "symbol=LTCBTC&side=BUY&type=LIMIT&price=0.01&quantity=1&timeInForce=GTC")
'{} -> test orders return empty JSON

End Sub

Function PublicBinance(Method As String, Optional MethodOptions As String) As String

'https://binance.com/home/api
Dim Url As String
PublicApiSite = "https://api.binance.com"
urlPath = "/api/v1/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicBinance = GetDataFromURL(Url, "GET")

End Function
Function PrivateBinance(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
Dim TimeCorrection As Long
'https://binance.com/home/api

'Get a 13-digit Nonce -> use the GetBinanceTime() to avoid a time correction
NonceUnique = GetBinanceTime() + 1000
TradeApiSite = "https://api.binance.com/api/v3/"

postdata = MethodOptions & "&timestamp=" & NonceUnique
APIsign = ComputeHash_C("SHA256", postdata, secretkey, "STRHEX")
Url = TradeApiSite & Method & "?" & postdata & "&signature=" & APIsign

'Binance requires a POST for orders, other commands are GETs
If InStr(Method, "order") > 0 Then
    HTTPMethod = "POST"
Else
    HTTPMethod = "GET"
End If

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open HTTPMethod, Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "X-MBX-APIKEY", apikey
objHTTP.Send get_url

objHTTP.WaitForResponse
PrivateBinance = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

Function GetBinanceTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicBinance time
JsonResponse = PublicBinance("time", "")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetBinanceTime = Json("serverTime")

Set Json = Nothing

End Function

