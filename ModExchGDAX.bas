Attribute VB_Name = "ModExchGDAX"
Sub TestGDAX()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String
Dim Passphrase As String

apiKey = "your api key here"
secretKey = "your secret key here"
Passphrase = "your passphrase here"

'Remove these 3 lines, unless you define 3 constants somewhere ( Public Const secretkey_gdax = "the key to use everywhere" etc )
apiKey = apikey_gdax
secretKey = secretkey_gdax
Passphrase = passphrase_gdax

Debug.Print PublicGDAX("time", "")
'{"iso":"2018-01-10T14:24:13.611Z","epoch":1515594253.611}
Debug.Print PublicGDAX("products", "/BTC-USD/book?level=2")
'{"sequence":4828391887,"bids":[["13779.97","8.44084168",17],["13774.46","0.0003",1],["13772.97","0.3513",1],["13759.8","0.00732578",1],["13755.01","0.00732578",1],["13755","0.2",1],["13754.55","0.00732578",1],["13754.54","9.32",1],["13750.03","27.02426867",1],["13750","0.0363901",1],["13749.99","2.41337397",12],["13749.76","0.0101",2],["13745.49","9.9318",1],["13745","0.015",1],["13744","0.025",1],["13743.88","0.00150364",1],["13743","0.1",1],["13741.79","0.9",1],["13741.13","0.00150364",1],["13740.15","9.2",1],["13740","0.01300726",23],["13738.38","0.0015",1],["13737","0.0051",1],["13736.54","0.025",1],["13736.44","0.00108181",1],["13735.99","0.01",1],["13735.97","0.00072437",1],["13735.63","0.0015",1],["13735.18","0.01",1],["13734.98","0.9",1]

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateGDAX("accounts", "GET", apiKey, secretKey, Passphrase)
'[{"id":"8a06fcff-f233-4b2a-b333-ec2ccd727956","currency":"BTC","balance":"0.0000000000000000","available":"0 etc...
Debug.Print PrivateGDAX("orders", "DELETE", apiKey, secretKey, Passphrase, "?product_id=BTC-USD")
'[]

End Sub

Function PublicGDAX(Method As String, Optional MethodOptions As String) As String

'https://docs.gdax.com/?php#api
Dim Url As String
PublicApiSite = "https://api.gdax.com"
urlPath = "/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicGDAX = WebRequestURL(Url, "GET")

End Function
Function PrivateGDAX(Method As String, HTTPMethod As String, apiKey As String, secretKey As String, Passphrase As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'https://docs.gdax.com/?php#api

'Get a 10-digit Nonce
NonceUnique = GetGDAXTime
TradeApiSite = "https://api.gdax.com"

SignMsg = NonceUnique & UCase(HTTPMethod) & "/" & Method & ""
APIsign = Base64Encode(ComputeHash_C("SHA256", SignMsg, Base64Decode(secretKey), "RAW"))

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open UCase(HTTPMethod), TradeApiSite & "/" & Method, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "CB-ACCESS-KEY", apiKey
objHTTP.setRequestHeader "CB-ACCESS-SIGN", APIsign
objHTTP.setRequestHeader "CB-ACCESS-TIMESTAMP", NonceUnique
objHTTP.setRequestHeader "CB-ACCESS-PASSPHRASE", Passphrase
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateGDAX = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

Function GetGDAXTime() As Double

Dim JsonResponse As String
Dim json As Object

'PublicGDAX time
JsonResponse = PublicGDAX("time", "")
Set json = JsonConverter.ParseJson(JsonResponse)
GetGDAXTime = Int(json("epoch"))
If GetGDAXTime = 0 Then
    TimeCorrection = -3600
    GetGDAXTime = CreateNonce(10)
    GetGDAXTime = Trim(Str((Val(GetGDAXTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set json = Nothing

End Function

