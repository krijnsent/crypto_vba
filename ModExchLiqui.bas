Attribute VB_Name = "ModExchKucoin"
Sub TestKucoin()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Kucoin will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretkey As String

apiKey = "your api key here"
secretkey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Kucoin = "the key to use everywhere" etc )
apiKey = apikey_kucoin
secretkey = secretkey_kucoin

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchKucoin"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestKucoinPublic")

'Testing error catching and replies
TestResult = PublicKucoin("open/tick", "")
'Debug.Print TestResult
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516555225011,"data":[{"coinType":"KCS","trading":true,"symbol":"KCS-BCH","lastDealPrice":0.0055,"buy":0.005425,"sell":0.0055,"change":-0.00014795,"coinTypePair":"BCH","sort":0,"feeRate":0.001,"volValue":90.38840317,"high":0.0059999,"datetime":1516555216000,"vol":16009.2128,"low":0.0053999,"changeRate":-0.0262},{"coinType":"KCS","trading":true,"sym etc...Test.IsEqual Left(TestResult, 40), "{""success"":true,""message"":"""",""result"":[{"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("msg"), "Operation succeeded."
Test.IsOk JsonResult("data")(1)("volValue") > 0
Test.IsOk JsonResult("data")(1)("lastDealPrice") > 0

TestResult = PublicKucoin("open/orders-buy", "?symbol=kcs-btc")
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516555225897,"data":[[8.3879E-4,50 etc...
'Debug.Print TestResult
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("msg"), "Operation succeeded."
Test.IsOk JsonResult("data")(1)(1) > 0

TestResult = GetKucoinTime()
'{}
Test.IsOk TestResult > 1500000000000#
Test.IsOk TestResult < 1600000000000#


'Unix time period:
Set Test = Suite.Test("TestKucoinPrivate")
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

TestResult = PrivateKucoin("user/info", apiKey, secretkey)
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516564519087,"data":{"referrer_code":"", etc...
'Debug.Print TestResult
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("msg"), "Operation succeeded."
Test.NotUndefined JsonResult("data")("referrer_code")


TestResult = PrivateKucoin("account/TFL/wallet/records", apiKey, secretkey, "type=DEPOSIT")
'{"success":true,"code":"OK","msg":"Operation succeeded.","timestamp":1516564519402,"data":{"total":1,"firstPage":true,"lastPage":false,"datas":[{"coinType":" etc...
'Debug.Print TestResult
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("success"), True
Test.IsEqual JsonResult("msg"), "Operation succeeded."
Test.IsOk JsonResult("data")("total") > 0


End Sub

Function PublicKucoin(Method As String, Optional MethodOptions As String) As String

'https://kucoinapidocs.docs.apiary.io/
Dim Url As String
PublicApiSite = "https://api.kucoin.com"
urlPath = "/v1/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicKucoin = WebRequestURL(Url, "GET")

End Function
Function PrivateKucoin(Method As String, apiKey As String, secretkey As String, Optional MethodOptions As String) As String

'https://kucoinapidocs.docs.apiary.io/
Dim NonceUnique As String
Dim Url As String
Dim postdata As String

'Kucoin wants a 13-digit Nonce, use time correction if needed
NonceUnique = GetKucoinTime()

'Arrange the MethodOptions parameters in ascending alphabetical order (lower cases first), then combine them with & (don't urlencode them, don't add ?, don't add extra &), e.g. amount=10&price=1.1&type=BUY
TradeApiSite = "https://api.kucoin.com"
ApiEndpoint = "/v1/" & Method
ApiForSign = ApiEndpoint & "/" & NonceUnique & "/" & MethodOptions
Base64ForSign = Base64Encode(ApiForSign)
APIsign = ComputeHash_C("SHA256", Base64ForSign, secretkey, "STRHEX")

Url = TradeApiSite & ApiEndpoint & "?" & MethodOptions

Dim headerDict As New Dictionary
headerDict.Add "KC-API-SIGNATURE", APIsign
headerDict.Add "KC-API-KEY", apiKey
headerDict.Add "KC-API-NONCE", NonceUnique
headerDict.Add "Content-Type", "application/json"

PrivateKucoin = WebRequestURL(Url, "GET", headerDict, postdata)

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
