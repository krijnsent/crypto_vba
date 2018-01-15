Attribute VB_Name = "ModExchHitBTC"
Sub TestHitBTC()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'HitBTC will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_HitBTC = "the key to use everywhere" etc )
apikey = apikey_hitbtc
secretkey = secretkey_hitbtc

Debug.Print PublicHitBTC("time")
'Example: {"timestamp":1516023792943}
Debug.Print PublicHitBTC("ticker", "BTCUSD/")
'{"ask":"14199.91","bid":"14191.39","last":"14199.98","low":"12900.00","high":"14200.00","open":"13382.15", etc..

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateHitBTC("balance", apikey, secretkey)
'{"balance":[{"currency_code":"1ST","cash":"0","reserved":"0"},{"currency_code":"8BT","cash":"0","reserved":"0"}, etc...
Debug.Print PrivateHitBTC("cancel_orders", apikey, secretkey, "&symbol=BTCUSD")
'{"ExecutionReport":[]} -> or a list of all cancelled trade numbers

End Sub

Function PublicHitBTC(Method As String, Optional MethodOptions As String) As String

'https://api.hitbtc.com/api/2/explore/
Dim Url As String
PublicApiSite = "https://api.hitbtc.com"
urlPath = "/api/1/public/" & MethodOptions & Method
Url = PublicApiSite & urlPath

PublicHitBTC = GetDataFromURL(Url, "GET")

End Function
Function PrivateHitBTC(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

'https://github.com/hitbtc-com/hitbtc-api#rest-api-reference

Dim NonceUnique As String
Dim postdata As String

'HitBTC nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)

TradeApiSite = "http://api.hitbtc.com"
urlPath = "/api/1/trading/" & Method & "?nonce=" & NonceUnique & "&apikey=" & apikey
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

