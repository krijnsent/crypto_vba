Attribute VB_Name = "ModExchKucoin"
Sub TestKucoin()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Kucoin will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apikey As String
Dim secretkey As String

apikey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Kucoin = "the key to use everywhere" etc )
apikey = apikey_kucoin
secretkey = secretkey_kucoin

Debug.Print PublicKucoin("open/tick", "")
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516555225011,"data":[{"coinType":"KCS","trading":true,"symbol":"KCS-BCH","lastDealPrice":0.0055,"buy":0.005425,"sell":0.0055,"change":-0.00014795,"coinTypePair":"BCH","sort":0,"feeRate":0.001,"volValue":90.38840317,"high":0.0059999,"datetime":1516555216000,"vol":16009.2128,"low":0.0053999,"changeRate":-0.0262},{"coinType":"KCS","trading":true,"sym etc...
Debug.Print PublicKucoin("open/orders-buy", "?symbol=kcs-btc")
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516555225897,"data":[[8.3879E-4,50 etc...
Debug.Print GetKucoinTime()
'{}

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateKucoin("user/info", apikey, secretkey)
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516564519087,"data":{"referrer_code":"", etc...
Debug.Print PrivateKucoin("account/TFL/wallet/records", apikey, secretkey, "type=DEPOSIT")
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516564519402,"data":{"total":1,"firstPage":true,"lastPage":false,"datas":[{"coinType":" etc...

End Sub

Function PublicKucoin(Method As String, Optional MethodOptions As String) As String

'https://kucoinapidocs.docs.apiary.io/
Dim Url As String
PublicApiSite = "https://api.kucoin.com"
urlPath = "/v1/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicKucoin = GetDataFromURL(Url, "GET")

End Function
Function PrivateKucoin(Method As String, apikey As String, secretkey As String, Optional MethodOptions As String) As String

'https://kucoinapidocs.docs.apiary.io/
Dim NonceUnique As String

'Kucoin wants a 13-digit Nonce, use time correction if needed
NonceUnique = GetKucoinTime()

'Arrange the MethodOptions parameters in ascending alphabetical order (lower cases first), then combine them with & (don't urlencode them, don't add ?, don't add extra &), e.g. amount=10&price=1.1&type=BUY
TradeApiSite = "https://api.kucoin.com"
ApiEndpoint = "/v1/" & Method
ApiForSign = ApiEndpoint & "/" & NonceUnique & "/" & MethodOptions
Base64ForSign = Base64Encode(ApiForSign)

APIsign = ComputeHash_C("SHA256", Base64ForSign, secretkey, "STRHEX")
'Debug.Print ApiEndpoint
'Debug.Print ApiForSign
'Debug.Print Base64ForSign
'Debug.Print APIsign

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open "GET", TradeApiSite & ApiEndpoint & "?" & MethodOptions, False
objHTTP.setRequestHeader "KC-API-SIGNATURE", APIsign
objHTTP.setRequestHeader "KC-API-KEY", apikey
objHTTP.setRequestHeader "KC-API-NONCE", NonceUnique
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateKucoin = objHTTP.ResponseText
Set objHTTP = Nothing

End Function

Function GetKucoinTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicKucoin time
JsonResponse = PublicKucoin("open/tick", "")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetKucoinTime = Json("timestamp")
If GetKucoinTime = 0 Then
    TimeCorrection = -3600
    GetKucoinTime = DateDiff("s", "1/1/1970", Now)
    GetKucoinTime = Trim(Str((Val(GetKucoinTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set Json = Nothing

End Function
