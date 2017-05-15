Attribute VB_Name = "ModExchBTCe"
Sub TestBTCe()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

'Testcase to be...

End Sub

Function PublicBTCe(Method As String, Optional MethodOptions As String) As String


PublicBTCe = ""

End Function
Function PrivateBTCe(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String

'BTC-e wants a 10-digit Nonce
NonceUnique = DateDiff("s", "1/1/1970", Now)
TradeApiSite = "https://btc-e.com/tapi/"

postdata = "method=" & Method & MethodOptions & "&nonce=" & NonceUnique
APIsign = ComputeHash_C("SHA512", postdata, secretkey, "STRHEX")

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", TradeApiSite, False
objHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objHTTP.SetRequestHeader "Key", apikey
objHTTP.SetRequestHeader "Sign", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateBTCe = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
