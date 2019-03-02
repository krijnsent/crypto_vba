Attribute VB_Name = "ModExchCoinspot"
Sub TestCoinspot()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_coinspot
secretKey = secretkey_coinspot

Debug.Print PublicCoinspot("latest", "")
'{"status":"ok","prices":{"btc":{"bid":"23000","ask":"23888.86","last":"23200"},"ltc":{"bid":"438","ask":"469.98","last":"440"},"doge":{"bid":"0.00700001","ask":"0.0089","last":"0.008"}}}

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateCoinspot("my/balances", apiKey, secretKey)
'{"status":"invalid"} / {"status":"no nonce"}
Debug.Print PrivateCoinspot("orders/history", apiKey, secretKey, "&cointype=LTC")
'{"status":"invalid"} / {"status":"no nonce"}
'ERROR: https://stackoverflow.com/questions/47799323/coinspot-api-with-powershell

End Sub

Function PublicCoinspot(Method As String, Optional MethodOptions As String) As String

'https://Coinspot.com/home/api
Dim Url As String

PublicApiSite = "https://www.coinspot.com.au"
urlPath = "/pubapi/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicCoinspot = WebRequestURL(Url, "GET")

End Function
Function PrivateCoinspot(Method As String, apiKey As String, secretKey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'https://Coinspot.com/home/api

'Get a 10-digit Nonce
NonceUnique = DateDiff("s", "1/1/1970", Now) & "0000000"
TradeApiSite = "https://www.coinspot.com.au"

postpath = "/api/" & Method
postdata = "nonce=" & NonceUnique & MethodOptions
APIsign = ComputeHash_C("SHA512", postdata, secretKey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'Debug.Print "POST: " & TradeApiSite & postdata
objHTTP.Open "POST", TradeApiSite & postpath, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "key", apiKey
objHTTP.setRequestHeader "sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateCoinspot = objHTTP.responseText
Set objHTTP = Nothing

End Function


