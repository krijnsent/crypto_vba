Attribute VB_Name = "ModExchLiqui"
Sub TestLiqui()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Liqui = "the key to use everywhere" etc )
apikey = apikey_liqui
secretkey = secretkey_liqui

Debug.Print PublicLiqui("info")
'Example: {"server_time":1508229311,"pairs":{"ltc_btc":{"decimal_places"... etc
Debug.Print PublicLiqui("ticker", "/eth_btc-san_btc")
'Example: {"eth_btc":{"high":0.06091721,"low":0.05726506,"avg":0.059091135,"vol":149.8217305710855367,"vol_cur":2549.93369699, etc

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateLiqui("getInfo", apikey, secretkey)
'{"success":1,"return":{"funds":{"btc":0.0,"ltc":0.0,"steem":0.0,"sbd":0.0,"dash":0.0, etc...
Debug.Print PrivateLiqui("TradeHistory", apikey, secretkey, "&since=" & t1 & "&end=" & t2 & "&")
'{"success":1,"return":{},"stat":{"isSuccess":true,"serverTime":"00:00:00.0008703","time":"0 etc...


End Sub

Function PublicLiqui(Method As String, Optional MethodOptions As String) As String

'https://www.Liqui.com/en-us/help/api#public-market-data

PublicApiSite = "https://api.liqui.io"
urlPath = "/api/3/" & Method & MethodOptions
Url = PublicApiSite & urlPath

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", Url
objHTTP.Send
objHTTP.WaitForResponse
PublicLiqui = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
Function PrivateLiqui(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

'https://www.Liqui.com/help/api#private-user-data

Dim NonceUnique As String
Dim postdata As String

'Liqui nonce: 16 characters
NonceUnique = DateDiff("s", "1/1/1970", Now)
'NonceUnique = NonceUnique & Right(Timer * 100, 2)

TradeApiSite = "https://api.liqui.io"
urlPath = "/tapi/"
postdata = "method=" & Method & MethodOptions & "&nonce=" & NonceUnique
Url = TradeApiSite & urlPath
APIsign = ComputeHash_C("SHA512", postdata, secretkey, "STRHEX")
Debug.Print postdata
Debug.Print Url
Debug.Print apikey
Debug.Print APIsign

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.SetRequestHeader "Key", apikey
objHTTP.SetRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateLiqui = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

