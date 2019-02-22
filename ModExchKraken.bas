Attribute VB_Name = "ModExchKraken"
Sub TestKraken()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Kraken will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_kraken = "the key to use everywhere" etc )
apiKey = apikey_kraken
secretKey = secretkey_kraken

Debug.Print PublicKraken("Time")
'Example: {"error":[],"result":{"unixtime":1494849819,"rfc1123":"Mon, 15 May 17 12:03:39 +0000"}}
Debug.Print PublicKraken("OHLC", "?pair=XXBTZEUR")
'{"error":[],"result":{"XXBTZEUR":[[1494806880,"1641.101","1642.850","1641.101"," etc...

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateKraken("Balance", apiKey, secretKey)
'{"error":[],"result":{"ZEUR":"15.35","KFEE":"935","XXBT": etc...
Debug.Print PrivateKraken("TradesHistory", apiKey, secretKey, "start=" & t1 & "&end=" & t2 & "&")
'{"error":[],"result":{"trades":{"TBSI6I-EO4KN-MLU4AI":{"ordertxid":"O7AERY-NCNDR-6WKLMU","pair":"XXMRZEUR","time":1493715960.4854,"type":"buy","ordertype":"limit","price": etc...


End Sub

Function PublicKraken(Method As String, Optional MethodOptions As String) As String

'https://www.kraken.com/en-us/help/api#public-market-data
Dim Url As String
PublicApiSite = "https://api.kraken.com"
urlPath = "/0/public/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicKraken = WebRequestURL(Url, "GET")

End Function
Function PrivateKraken(Method As String, apiKey As String, secretKey As String, Optional MethodOptions As String) As String

'https://www.kraken.com/help/api#private-user-data

Dim NonceUnique As String
Dim postdata As String

'Kraken nonce: 16 characters
NonceUnique = CreateNonce(16)

TradeApiSite = "https://api.kraken.com"
urlPath = "/0/private/" & Method
postdata = MethodOptions & "nonce=" & NonceUnique

Url = TradeApiSite & urlPath
APIsign = ComputeHash_C("SHA512", urlPath & ComputeHash_C("SHA256", NonceUnique & postdata, "", "RAW"), Base64Decode(secretKey), "STR64")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.setRequestHeader "API-Key", apiKey
objHTTP.setRequestHeader "API-Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateKraken = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
