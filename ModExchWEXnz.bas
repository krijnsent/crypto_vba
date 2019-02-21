Attribute VB_Name = "ModExchWEXnz"
Sub TestWEXnz()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'WEXnz will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretkey As String

apiKey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_WEXnz = "the key to use everywhere" etc )
apiKey = apikey_wexnz
secretkey = secretkey_wexnz

Debug.Print PublicWEXnz("depth", "/btc_eur")
'{"btc_eur":{"asks":[[1615,1.69215502],[1615.93,0.13653712],[1615.95753,0.00989219],
Debug.Print PublicWEXnz("ticker", "/ltc_btc-btc_eur")
'{"ltc_btc":{"high":0.01453,"low":0.012,"avg":0.013265,"vol":3266.44821,"vol_cur":247301.29778,"last":0.0143,"buy":0.01432,"sell":0.0143,"updated":1495026094},"btc_eur":{"high": etc...

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

'Debug.Print t1, t2
Debug.Print PrivateWEXnz("getInfo", apiKey, secretkey)
'{"success":1,"return":{"funds":{"usd":0,"btc":0.14,"ltc":0,"nmc":0, etc...
Debug.Print PrivateWEXnz("TradeHistory", apiKey, secretkey, "&since=" & t1 & "&end=" & t2)
'{"success":1,"return":{"101927904":{"pair":"btc_eur","type":"sell","amount":0.01061285,"rate":1509 etc...

End Sub

Function PublicWEXnz(Method As String, Optional MethodOptions As String) As String

'https://wex.nz/api/3/docs
Dim Url As String
PublicApiSite = "https://wex.nz"
urlPath = "/api/3/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicWEXnz = WebRequestURL(Url, "GET")

End Function
Function PrivateWEXnz(Method As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

'https://wex.nz/tapi/docs
Dim NonceUnique As String

'BTC-e wants a 10-digit Nonce
NonceUnique = CreateNonce(10)
TradeApiSite = "https://wex.nz/tapi/"

postdata = "method=" & Method & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", postdata, secretkey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", TradeApiSite, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "Key", apiKey
objHTTP.setRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateWEXnz = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

