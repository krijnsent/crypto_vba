Attribute VB_Name = "ModExchCoinone"
Sub TestCoinone()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretkey As String

apiKey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_coinone
secretkey = secretkey_coinone

Debug.Print PublicCoinone("ticker", "")
'{"result":"success","volume":"5448.8622","last":"8135000","yesterday_last":"8133500","timestamp":"1510072143","yesterday_low":"8026000","high":"8330000","currency":"btc","low":"8026000","errorCode":"0","yesterday_first":"8297000","yesterday_volume":"5809.1167","yesterday_high":"8330000","first":"8230000"}
Debug.Print PublicCoinone("trades", "?currency=btc&period=hour")
'{"errorCode":"0","timestamp":"1510072144","completeOrders":[{"timestamp":"1510068569","price":"8153000","qty":"0.1722"},{"timestamp":"1510068576","price":"8153000","qty":"0.0765"},{"timestamp":"1510068669","price":"8149000","qty":"0.0764"},{"timestamp":"1510068687","price":"8155000","qty":"0.2067"},{"timestamp":"1510068687","price":"8155000","qty":"0.8763"},{"timestamp":"1510068688","price":"8155500","qty":"0.1476"}, etc.

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateCoinone("account/balance", apiKey, secretkey)
'{"errorCode":"0","result":"success","btc":{"avail":"0.00000000","balance":"0.00000000"},"normalWallets":[],"bch":{"avail":"0.00000000","balance":"0.00000000"},"qtum":{"avail":"0.00000000","balance":"0.00000000"},"krw":{"avail":"0","balance":"0"},"ltc":{"avail":"0.00000000","balance":"0.00000000"},"etc":{"avail":"0.00000000","balance":"0.00000000"},"eth":{"avail":"0.00000000","balance":"0.00000000"},"xrp":{"avail":"0.00000000","balance":"0.00000000"}} etc...
Debug.Print PrivateCoinone("order/complete_orders", apiKey, secretkey, "&currency=eth")
'{"errorCode":"0","completeOrders":[],"result":"success"}

End Sub

Function PublicCoinone(Method As String, Optional MethodOptions As String) As String

'https://Coinone.com/home/api
Dim Url As String
PublicApiSite = "https://api.coinone.co.kr/"
urlPath = Method & "/" & MethodOptions
Url = PublicApiSite & urlPath

PublicCoinone = WebRequestURL(Url, "GET")

End Function
Function PrivateCoinone(Method As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'http://doc.coinone.co.kr/

'Get a 14-digit Nonce
NonceUnique = CreateNonce(14)
'NonceUnique = "1510140617707865"
TradeApiSite = "https://api.coinone.co.kr/v2/"

Url = TradeApiSite & Method

postdata = "access_token=" & apiKey & "&" & "nonce=" & NonceUnique
If MethodOptions <> "" Then
    postdata = postdata & MethodOptions
End If

postdata_json_txt = Replace(postdata, "=", Chr(34) & ":" & Chr(34))
postdata_json_txt = Replace(postdata_json_txt, "&", Chr(34) & "," & Chr(34))
postdata_json_txt = "{" & Chr(34) & postdata_json_txt & Chr(34) & "}"
postdata64 = Base64Encode(postdata_json_txt)

APIsign = ComputeHash_C("SHA512", Base64Encode(postdata_json_txt), secretkey, "STRHEX")

'' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.setRequestHeader "X-COINONE-PAYLOAD", postdata64
objHTTP.setRequestHeader "X-COINONE-SIGNATURE", APIsign
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateCoinone = objHTTP.ResponseText
Set objHTTP = Nothing

End Function
