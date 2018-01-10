Attribute VB_Name = "ModExchGDAX"
Sub TestGDAX()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String
Dim passphrase As String

apikey = "your api key here"
secretkey = "your secret key here"
passphrase = "your passphrase here"

'Remove these 3 lines, unless you define 3 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apikey = apikey_gdax
secretkey = secretkey_gdax
passphrase = passphrase_gdax

Debug.Print PublicGDAX("time", "")
'{"iso":"2018-01-10T14:24:13.611Z","epoch":1515594253.611}
Debug.Print PublicGDAX("products", "/BTC-USD/book?level=2")
'{"sequence":4828391887,"bids":[["13779.97","8.44084168",17],["13774.46","0.0003",1],["13772.97","0.3513",1],["13759.8","0.00732578",1],["13755.01","0.00732578",1],["13755","0.2",1],["13754.55","0.00732578",1],["13754.54","9.32",1],["13750.03","27.02426867",1],["13750","0.0363901",1],["13749.99","2.41337397",12],["13749.76","0.0101",2],["13745.49","9.9318",1],["13745","0.015",1],["13744","0.025",1],["13743.88","0.00150364",1],["13743","0.1",1],["13741.79","0.9",1],["13741.13","0.00150364",1],["13740.15","9.2",1],["13740","0.01300726",23],["13738.38","0.0015",1],["13737","0.0051",1],["13736.54","0.025",1],["13736.44","0.00108181",1],["13735.99","0.01",1],["13735.97","0.00072437",1],["13735.63","0.0015",1],["13735.18","0.01",1],["13734.98","0.9",1]

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateGDAX("accounts", apikey, secretkey, passphrase)
'{"success":true,"message":"","result":[{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"},{"Currency":"XMR","Balance":0.00000000,"Available":0.00000000,"Pending":0.00000000,"CryptoAddress":etc...
'Debug.Print PrivateGDAX("account/getbalance", apikey, secretkey, "&currency=ETH")
'{"success":true,"message":"","result":{"Currency":"BTC","Balance":1.65740000,"Available":1.65740000,"Pending":0.00000000,"CryptoAddress":"1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa"}}

End Sub

Function PublicGDAX(Method As String, Optional MethodOptions As String) As String

'https://GDAX.com/home/api
Dim Url As String
PublicApiSite = "https://api.gdax.com"
urlPath = "/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicGDAX = GetDataFromURL(Url, "GET")

End Function
Function PrivateGDAX(Method As String, apikey As String, secretkey As String, passphrase As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'https://GDAX.com/home/api

'Get a 10-digit Nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)
TradeApiSite = "https://api.gdax.com"

'Under development, no test account available, so no
'postdata = Method & "?apikey=" & apikey & MethodOptions & "&nonce=" & NonceUnique
'APIsign = ComputeHash_C("SHA256", TradeApiSite & postdata, secretkey, "STRHEX")
APIsign = "NOTHING_HERE_YET"

'All REST requests must contain the following headers:
'CB-ACCESS-KEY The api key as a string.
'CB-ACCESS-SIGN The base64-encoded signature (see Signing a Message).
'CB-ACCESS-TIMESTAMP A timestamp for your request.
'CB-ACCESS-PASSPHRASE The passphrase you specified when creating the API key.
'All request bodies should have content type application/json and be valid JSON.
'The CB-ACCESS-SIGN header is generated by creating a sha256 HMAC using the base64-decoded secret key on the prehash string timestamp + method + requestPath + body (where + represents string concatenation) and base64-encode the output. The timestamp value is the same as the CB-ACCESS-TIMESTAMP header.
'The body is the request body string or omitted if there is no request body (typically for GET requests).
'The method should be UPPER CASE.

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'Debug.Print "POST: " & TradeApiSite & postdata
objHTTP.Open "GET", TradeApiSite & postdata, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "CB-ACCESS-KEY", apikey
objHTTP.setRequestHeader "CB-ACCESS-SIGN", APIsign
objHTTP.setRequestHeader "CB-ACCESS-TIMESTAMP", NonceUnique
objHTTP.setRequestHeader "CB-ACCESS-PASSPHRASE", passphrase
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateGDAX = objHTTP.ResponseText
Set objHTTP = Nothing

End Function


