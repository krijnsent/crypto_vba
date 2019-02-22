Attribute VB_Name = "ModExchKucoin"
Sub TestKucoin()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Kucoin will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Kucoin = "the key to use everywhere" etc )
apiKey = apikey_kucoin
secretKey = secretkey_kucoin

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchKucoin"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestKucoinPublic")

'Error, unknown command
TestResult = PublicKucoin("AnUnknownCommand")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, missing parameters
TestResult = PublicKucoin("market/orderbook/level1")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400


TestResult = PublicKucoin("market/allTickers", "")
'Debug.Print TestResult
'{"code":"200000","data":{"ticker":[{"symbol":"LOOM-BTC","high":"0.00001204","vol":"39738.31683935","last":"0.00001187","low":"0.00001151","buy":"0.00001172","sell":"0.00001187","changePrice":"0.00000025","changeRate":"0.0215"},etc...
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "changePrice") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data")("ticker").Count > 100
Test.IsOk Len(JsonResult("data")("ticker")(9)("symbol")) > 0
Test.IsOk JsonResult("data")("ticker")(3)("vol") > 0

TestResult = PublicKucoin("market/orderbook/level2_20", "?symbol=KCS-BTC")
'Debug.Print TestResult
'{"code":"200000","data":{"sequence":"1550467431550","asks":[["0.00011794","184.4706"],["0.00011795","48.7387"],["0.00011796","154.9647"],
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "sequence") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data")("time") > 1500000000000#
Test.IsEqual JsonResult("data")("asks").Count, 20
Test.IsEqual JsonResult("data")("bids").Count, 20
Test.IsOk JsonResult("data")("asks")(1)(1) > 0
Test.IsOk JsonResult("data")("asks")(1)(2) > 0

' Create a new test
Set Test = Suite.Test("TestKucoinTime")
TestResult = GetKucoinTime()
Test.IsOk TestResult > 1500000000000#
Test.IsOk TestResult < 1600000000000#


Set Test = Suite.Test("TestKucoinPrivate")

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

TestResult = PrivateKucoin("accounts", "GET", apiKey, secretKey, passphrase_kucoin)
'Debug.Print TestResult
'{"code":"200000","data":[{"balance":"15.827819","available":"15.827819","holds":"0","currency":"KCS","id":"5c6a4a1d81a34e1da97","type":"trade"},{"balance":"2.12058951","available":"2.12058951",", etc...
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data").Count > 20
Test.IsOk JsonResult("data")(1)("balance") > 0

'Get only KCS account amount
Dim OptDict As New Dictionary
OptDict.Add "currency", "KCS"
TestResult = PrivateKucoin("accounts", "GET", apiKey, secretKey, passphrase_kucoin, OptDict)
'Debug.Print TestResult
'{"code":"200000","data":[{"balance":"15.82887819","available":"15.82887819","holds":"0","currency":"KCS","id":"5c6a4a1d81a34e1da97","type":"trade"}]}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsEqual JsonResult("data").Count, 1
Test.IsOk JsonResult("data")(1)("balance") > 0

'Create a main LTC account (if it doesn't exist)
Dim OptDict2 As New Dictionary
OptDict2.Add "currency", "LTC"
TestResult = PrivateKucoin("accounts", "POST", apiKey, secretKey, passphrase_kucoin, OptDict2)
'Debug.Print TestResult
'{"code":"400100","msg":"type can not be empty"}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 400100
Test.IsEqual JsonResult("msg"), "type can not be empty"

OptDict2.Add "type", "main"
TestResult = PrivateKucoin("accounts", "POST", apiKey, secretKey, passphrase_kucoin, OptDict2)
'Debug.Print TestResult
'FIRST TIME RESULT: {"code":"200000","data":{"id":"5c7556e3cbfc7b24adada1a9"}}
'NEXT RESULT: {"code":"230005","msg":"account already exists"}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "msg") + InStr(TestResult, "data") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 230005


End Sub

Function PublicKucoin(Method As String, Optional MethodOptions As String) As String

'https://docs.kucoin.com/
Dim Url As String
PublicApiSite = "https://openapi-v2.kucoin.com/api"
urlPath = "/v1/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicKucoin = WebRequestURL(Url, "GET")

End Function
Function PrivateKucoin(Method As String, ReqType As String, apiKey As String, secretKey As String, Passphrase As String, Optional OptionsDict As Dictionary) As String

'https://docs.kucoin.com/
Dim NonceUnique As String
Dim Url As String
Dim postdata As String

'Kucoin wants a 13-digit Nonce, use time correction if needed
NonceUnique = GetKucoinTime()

TradeApiSite = "https://openapi-v2.kucoin.com"
ApiEndPoint = "/api/v1/" & Method
'e.g. /api/v1/deposit-addresses?currency=BTC

If ReqType = "GET" Or ReqType = "DELETE" Then
    'For GET, DELETE request, all query parameters need to be included in the request url. (e.g. /api/v1/accounts?currency=BTC)
    MethodTxt = DictToString(OptionsDict, "URLENC")
    If MethodTxt <> "" Then ApiEndPoint = ApiEndPoint & "?" & MethodTxt
    ReqBody = ""
Else
    'For POST, PUT request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}). Do not include extra spaces in JSON strings.
    MethodTxt = ""
    ReqBody = DictToString(OptionsDict, "JSON")
    postdata = ReqBody
End If

ApiForSign = NonceUnique & ReqType & ApiEndPoint & ReqBody
APIsign = ComputeHash_C("SHA256", ApiForSign, secretKey, "STR64")

Url = TradeApiSite & ApiEndPoint

Dim headerDict As New Dictionary
headerDict.Add "KC-API-KEY", apiKey
headerDict.Add "KC-API-SIGN", APIsign
headerDict.Add "KC-API-TIMESTAMP", NonceUnique
headerDict.Add "KC-API-PASSPHRASE", Passphrase
headerDict.Add "Content-Type", "application/json"

PrivateKucoin = WebRequestURL(Url, ReqType, headerDict, postdata)

End Function

Function GetKucoinTime() As Double

Dim JsonResponse As String
Dim json As Object

'PublicKucoin time
JsonResponse = PublicKucoin("timestamp", "")
Set json = JsonConverter.ParseJson(JsonResponse)
GetKucoinTime = json("data")
If GetKucoinTime = 0 Then
    TimeCorrection = -3600
    GetKucoinTime = DateDiff("s", "1/1/1970", Now)
    GetKucoinTime = Trim(Str((Val(GetKucoinTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set json = Nothing

End Function
