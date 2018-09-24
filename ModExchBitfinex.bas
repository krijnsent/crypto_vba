Attribute VB_Name = "ModExchBitfinex"
Sub TestBitfinex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apikey = apikey_bitfinex
secretkey = secretkey_bitfinex

Debug.Print PublicBitfinex("symbols", "")
'["btcusd","ltcusd","ltcbtc","ethusd","ethbtc","etcbtc","etcusd","rrtusd"...
Debug.Print PublicBitfinex("pubticker", "ltcbtc")
'{"mid":"0.0171145","bid":"0.017113","ask":"0.017116","last_price":"0.017105","low":"0.01666","high":"0.01721","volume":"85227.17880718","timestamp":"1515663208.679153"}

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateBitfinex("balances", apikey, secretkey)

End Sub

Function PublicBitfinex(Method As String, Optional MethodOptions As String) As String

'https://bittrex.com/home/api
Dim Url As String
PublicApiSite = "https://api.bitfinex.com"
urlPath = "/v1/" & Method & "/" & MethodOptions
Url = PublicApiSite & urlPath

PublicBitfinex = WebRequestURL(Url, "GET")

End Function
Function PrivateBitfinex(Method As String, apikey As String, secretkey As String, Optional MethodOptions As Collection)

Dim NonceUnique As String
Dim Json As String
Dim PayloadDict As Scripting.Dictionary

NonceUnique = CreateNonce(15)
'see the general Bitfinex documentation here: https://bitfinex.readme.io/v1/docs/rest-general

'the payload has to look like this: payload = parameters-object -> JSON encode -> base64
'see the authenticated endpoints documentation here: https://bitfinex.readme.io/v1/docs/rest-auth
Set PayloadDict = New Dictionary
PayloadDict("request") = "/v1/" & Method
PayloadDict("nonce") = NonceUnique
If Not MethodOptions Is Nothing Then
    Set PayloadDict("options") = MethodOptions
End If

Json = Replace(ConvertToJson(PayloadDict), "/", "\/")
Payload = Base64Encode(Json)

'signature = HMAC-SHA384(payload, api-secret).digest('hex')
ApiSite = "https://api.bitfinex.com"
signature = ComputeHash_C("SHA384", Payload, secretkey, "STRHEX")

Url = ApiSite & "/v1/" & Method
HTTPMethod = "POST"

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open HTTPMethod, Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "X-BFX-APIKEY", apikey
objHTTP.setRequestHeader "X-BFX-PAYLOAD", Payload
objHTTP.setRequestHeader "X-BFX-SIGNATURE", signature
objHTTP.Send get_url

objHTTP.WaitForResponse
PrivateBitfinex = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
