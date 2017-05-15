Attribute VB_Name = "ModExchPoloniex"
Sub TestPoloniex()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA



End Sub

Function PublicPoloniex(Method As String, Optional MethodOptions As String) As String


PublicPoloniex = ""

End Function
Function PrivatePoloniex(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
Dim postdata As String

'Poloniex nonce: 16 characters
NonceUnique = DateDiff("s", "1/1/1970", Now)
NonceUnique = NonceUnique & Right(Timer * 100, 2) & "0000"

URL = "https://poloniex.com/tradingApi"
postdata = "command=" & Method & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", postdata, secretkey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", URL, False
objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.SetRequestHeader "Key", apikey
objHTTP.SetRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivatePoloniex = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
