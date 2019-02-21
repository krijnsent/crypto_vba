Attribute VB_Name = "ModExchHitBTC"
Sub TestHitBTC()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'HitBTC will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretkey As String

apiKey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_HitBTC = "the key to use everywhere" etc )
apiKey = apikey_hitbtc
secretkey = secretkey_hitbtc

Debug.Print PublicHitBTC("time")
'Example: {"timestamp":1516023792943}
Debug.Print PublicHitBTC("ticker", "BTCUSD/")
'{"ask":"14199.91","bid":"14191.39","last":"14199.98","low":"12900.00","high":"14200.00","open":"13382.15", etc..

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateHitBTC("balance", apiKey, secretkey)
'{"balance":[{"currency_code":"1ST","cash":"0","reserved":"0"},{"currency_code":"8BT","cash":"0","reserved":"0"}, etc...
Debug.Print PrivateHitBTC("cancel_orders", apiKey, secretkey, "&symbol=BTCUSD")
'{"ExecutionReport":[]} -> or a list of all cancelled trade numbers

Debug.Print PublicHitBTC2("symbol")
'[{"id":"BCNBTC","baseCurrency":"BCN","quoteCurrency":"BTC","quantityIncrement":"100","tickSize":"0.0000000001","takeLiquidityR etc...
Debug.Print PublicHitBTC2("ticker", "/BCHUSD")
'{"ask":"1228.13454","bid":"1225.62444","last":"1227.17775","open":"1269.12060","low":"1 etc...

Debug.Print PrivateHitBTC2("account/balance", "GET", apiKey, secretkey)
'[{"currency":"DOGE","available":"0.00000000","reserved":"0.00000000"},{" etc...
Debug.Print PrivateHitBTC2("history/trades", "GET", apiKey, secretkey, "?symbol=BTCUSD")
'e.g. []
Debug.Print PrivateHitBTC2("order", "DELETE", apiKey, secretkey, "?symbol=BTCUSD")
'e.g. []

End Sub

Function PublicHitBTC(Method As String, Optional MethodOptions As String) As String

'https://api.hitbtc.com/api/2/explore/
Dim Url As String
PublicApiSite = "https://api.hitbtc.com"
urlPath = "/api/1/public/" & MethodOptions & Method
Url = PublicApiSite & urlPath

PublicHitBTC = WebRequestURL(Url, "GET")

End Function
Function PrivateHitBTC(Method As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

'https://github.com/hitbtc-com/hitbtc-api#rest-api-reference

Dim NonceUnique As String
Dim postdata As String

'HitBTC nonce
NonceUnique = CreateNonce(10)

TradeApiSite = "http://api.hitbtc.com"
urlPath = "/api/1/trading/" & Method & "?nonce=" & NonceUnique & "&apikey=" & apiKey
postdata = MethodOptions

Url = TradeApiSite & urlPath
APIsign = LCase(ComputeHash_C("SHA512", urlPath & postdata, secretkey, "STRHEX"))

'HitBTC requires a POST for orders, other commands are GETs
If InStr(Method, "_order") > 0 Then
    HTTPMethod = "POST"
Else
    HTTPMethod = "GET"
End If

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open HTTPMethod, Url & postdata, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.setRequestHeader "X-Signature", APIsign
objHTTP.Send ("")

objHTTP.WaitForResponse
PrivateHitBTC = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

Function PublicHitBTC2(Method As String, Optional MethodOptions As String) As String

'https://api.hitbtc.com/api/2/explore/
Dim Url As String
PublicApiSite = "https://api.hitbtc.com"
urlPath = "/api/2/public/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicHitBTC2 = WebRequestURL(Url, "GET")

End Function

Function PrivateHitBTC2(Method As String, HTTPMethod As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

'https://api.hitbtc.com/api/2/explore/
'Authorisation: https://stackoverflow.com/questions/34637034/curl-u-equivalent-in-http-request

Dim NonceUnique As String
Dim postdata As String

'HitBTC nonce
NonceUnique = CreateNonce(10)

TradeApiSite = "http://api.hitbtc.com"
urlPath = "/api/2/" & Method & MethodOptions
Url = TradeApiSite & urlPath

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open HTTPMethod, Url, False
objHTTP.setRequestHeader "Accept", "application/json"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Authorization", "Basic " & Base64Encode(apiKey & ":" & secretkey)
objHTTP.Send ("")

objHTTP.WaitForResponse
PrivateHitBTC2 = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
