Attribute VB_Name = "ModExchCryptopia"
Sub TestCryptopia()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Cryptopia = "the key to use everywhere" etc )
apikey = apikey_cryptopia
secretkey = secretkey_cryptopia

Debug.Print PublicCryptopia("GetCurrencies")
'Example: {"Success":true,"Message":null,"Data":[{"Id":331,"Name":"1337","Symbol":"1337","Algorithm":"POS"... etc
Debug.Print PublicCryptopia("GetMarket", "/DOT_BTC")
'Example: {"Success":true,"Message":null,"Data":{"TradePairId":100,"Label":"DOT/BTC","AskPrice":0.00000085, etc

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateCryptopia("GetBalance", apikey, secretkey, "Currency=")
'{"Success":true,"Error":null,"Data":[{"CurrencyId":331,"Symbol":"1337","Total":0.00000000,"Available":0.00000000, etc...
Debug.Print PrivateCryptopia("GetTradeHistory", apikey, secretkey, "Market=DOT/BTC")
'{"Success":true,"Error":null,"Data":[ etc...

End Sub

Function PublicCryptopia(Method As String, Optional MethodOptions As String) As String

'https://www.cryptopia.co.nz/forum/Thread/255

PublicApiSite = "https://www.cryptopia.co.nz"
urlPath = "/api/" & Method & MethodOptions
Url = PublicApiSite & urlPath

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", Url
objHTTP.Send
objHTTP.WaitForResponse
PublicCryptopia = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
Function PrivateCryptopia(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

'https://www.cryptopia.co.nz/forum/Thread/256

Dim NonceUnique As String
Dim postdata As String
Dim Url As String

'Cryptopia nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)

TradeApiSite = "https://www.cryptopia.co.nz"
urlPath = "/api/"
Url = TradeApiSite & urlPath & Method
UrlEnc = LCase(URLEncode(Url))

postdata = MethodOptions '{"Currency":""}
postdataJsonTxt = Replace(postdata, "=", Chr(34) & ":" & Chr(34))
postdataJsonTxt = Replace(postdataJsonTxt, "&", Chr(34) & "," & Chr(34))
postdataJsonTxt = "{" & Chr(34) & postdataJsonTxt & Chr(34) & "}"
req64 = ComputeHash_C("MD5", postdataJsonTxt, "", "STR64")

Signature = apikey & "POST" & UrlEnc & NonceUnique & req64
hmacSignature = ComputeHash_C("SHA256", Signature, Base64Decode(secretkey), "STR64")
HeaderValue = "amx " & apikey & ":" & hmacSignature & ":" & NonceUnique

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.SetRequestHeader "Content-Type", "application/json"
objHTTP.SetRequestHeader "Authorization", HeaderValue
objHTTP.Send (postdataJsonTxt)

objHTTP.WaitForResponse
PrivateCryptopia = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
