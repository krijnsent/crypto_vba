Attribute VB_Name = "ModExchHitBTC"
Sub TestHitBTC()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'https://api.hitbtc.com/api/2/explore/
'https://github.com/hitbtc-com/hitbtc-api#rest-api-reference
'HitBTC will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_HitBTC = "the key to use everywhere" etc )
apiKey = apikey_hitbtc
secretKey = secretkey_hitbtc

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchHitBTC"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestHitBTCPublic v2")

'Error, unknown command
TestResult = PublicHitBTCv2("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0, "unknowncommand 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "unknowncommand 2 failed, result: ${1}"


'Error, wrong parameter
Dim Params As New Dictionary
Params.Add "symbol", "BLABLA"
TestResult = PublicHitBTCv2("trades", "GET", Params)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"timestamp":"2021-02-09T17:30:24.738+00:00","path":"/api/2/public/trades/BLABLA","status":400,"error":{"code":2001,"description":"Try get /public/symbol, to get list of all available symbols.","message":"No such symbol: BLABLA"},"requestId":"eecd7978-102065517"}}
Test.IsOk InStr(TestResult, "error") > 0, "trades params 1 failed, result: ${1}"
Test.IsOk InStr(TestResult, "No such symbol") > 0, "trades params 2 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400, "trades params 3 failed, result: ${1}"
Test.IsEqual JsonResult("response_txt")("error")("code"), 2001, "trades params 4 failed, result: ${1}"

'Simple request without parameters
TestResult = PublicHitBTCv2("currency", "GET")
'Example: [{"id":"DDF","fullName":"DDF","crypto":true,"payinEnabled":false,"payinPaymentId":false,"payinConfirmations":2,"payoutEnabled":true,"payoutIsPaymentId":false,"transferEnabled":true,"delisted":false,"payoutFee":"646"},{"id":"ZRX","fullName":"0x Protocol","crypto":true,"payinEnabled":true,"payinPaymentId":false,"payinConfirmations":2,"payoutEnabled":true,"payoutIsPaymentId":false,"transferEnabled":true,"delisted":false,"payoutFee":"26.45"},{"id":"ACO","fullName":"A!Coin","crypto":true etc...
Test.IsOk InStr(TestResult, "payoutFee") > 0, "currency 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count >= 100, "currency 2 failed, result: ${1}"
Test.IsOk Len(JsonResult(1)("id")) >= 3, "currency 3 failed, result: ${1}"

'Request with parameter
Dim Params2 As New Dictionary
Params2.Add "currency", "ETH"
TestResult = PublicHitBTCv2("currency", "GET", Params2)
'{"id":"ETH","fullName":"Ethereum","crypto":true,"payinEnabled":true,"payinPaymentId":false,"payinConfirmations":2,"payoutEnabled":true,"payoutIsPaymentId":false,"transferEnabled":true,"delisted":false,"payoutFee":"0.0428"}
Test.IsOk InStr(TestResult, "Ethereum") > 0, "currency params 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("id"), "ETH", "currency params 2 failed, result: ${1}"
Test.IsEqual JsonResult("crypto"), True, "currency params 3 failed, result: ${1}"
Test.IsEqual JsonResult("delisted"), False, "currency params 4 failed, result: ${1}"

'Request with parameters
Dim Params3 As New Dictionary
Params3.Add "symbol", "ETHBTC"
Params3.Add "sort", "ASC"
Params3.Add "limit", 10
TestResult = PublicHitBTCv2("trades", "GET", Params3)
'[{"id":3462311,"price":"0.006000","quantity":"0.001","side":"buy","timestamp":"2015-08-20T19:01:23.764Z"},{"id":3462314,"price":"0.006000","quantity":"0.001","side":"buy","timestamp":"2018-07-10T16:11:35.511Z"},etc...
Test.IsOk InStr(TestResult, "timestamp") > 0, "trades2 params 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult(1)("id") > 0, "trades2 params 2 failed, result: ${1}"
Test.IsOk Val(JsonResult(1)("quantity")) > 0, "trades2 params 3 failed, result: ${1}"
Test.IsEqual JsonResult(1)("side"), "buy", "trades2 params 4 failed, result: ${1}"

Set Test = Suite.Test("TestHitBTCPrivate v2")

TestResult = PrivateHitBTCv2("trading/balance", "GET", Cred)
'[{"currency":"1ST","available":"0","reserved":"0"},{"currency":"8BT","available":"0","reserved":"0"},{"currency":"ABA","available":"0","reserved":"0"},{"currency":"ABTC","available":"0","reserved":"0"},{"currency":"ABYSS","available":"0","reserved":"0"} etc...
Test.IsOk InStr(TestResult, "available") > 0
Test.IsOk InStr(TestResult, "reserved") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
'Loop through all coins
For Each Coin In JsonResult
    If Coin("available") + Coin("reserved") > 0 Then
        'Debug.Print Coin("currency"), Coin("available") + Coin("reserved")
        Test.IsOk Len(Coin("currency")) >= 3
    End If
Next Coin
Test.IsOk Len(JsonResult(1)("currency")) > 0
Test.IsOk Val(JsonResult(2)("available")) >= 0

Dim Params4 As New Dictionary
Params4.Add "symbol", "DOGEETH"
TestResult = PrivateHitBTCv2("history/trades", "GET", Cred, Params4)
'e.g. [{"id":215639995,"clientOrderId":"4ab37988ea9545aeb325fc60931fbaa3","orderId":19837911730,"symbol":"DOGEETH","side":"sell","quantity": etc.
If TestResult = "[]" Then
    Test.IsEqual TestResult, "[]"
Else
    Test.IsOk InStr(TestResult, "clientOrderId") > 0
    Test.IsOk InStr(TestResult, "symbol") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk Len(JsonResult(1)("symbol")) >= 6
    Test.IsOk JsonResult(2)("orderId") > 0
End If

'Delete all orders DOGE-ETH
Dim Params5 As New Dictionary
Params5.Add "symbol", "DOGEETH"
TestResult = PrivateHitBTCv2("order", "DELETE", Cred, Params5)
'e.g. [{"id": 0,"clientOrderId": "d8574207d9e3b16a4a5511753eeef175","symbol": "DOGEETH","side": "sell","status": "canceled","type": "limit", etc...
If InStr(TestResult, "NO VALID JSON RETURNED") > 0 Then
    Test.IsOk InStr(TestResult, ":200") > 0
Else
    If TestResult <> "[]" Then
        Test.IsOk InStr(TestResult, "clientOrderId") > 0
        Test.IsOk InStr(TestResult, "symbol") > 0
        Set JsonResult = JsonConverter.ParseJson(TestResult)
        Test.IsEqual JsonResult(1)("symbol"), "DOGEETH"
        Test.IsOk Len(JsonResult(1)("side")) >= 3
    End If
End If

'Create an order, but trigger an error
Dim Params6 As New Dictionary
Params6.Add "symbol", "ETHBTC"
Params6.Add "side", "sell"
Params6.Add "quantity", "0.000005"
Params6.Add "price", "1"
TestResult = PrivateHitBTCv2("order", "POST", Cred, Params6)
'e.g. {"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"error":{"code":20001,"message":"Insufficient funds","description":"Check that the funds are sufficient, given commissions"}}}
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"error":{"code":2011,"message":"Quantity too low","description":"Minimum quantity 0.0001"}}}
'if OK, e.g. {"id": 0,"clientOrderId": "d8574207d9e3b16a4a5511753eeef175","symbol": "ETHBTC","side": "sell","status": "new","type": "limit","timeInForce": "GTC","quantity": "0.063","price": "0.046016","cumQuantity": "0.000","postOnly": false,"createdAt": "2017-05-15T17:01:05.092Z","updatedAt": "2017-05-15T17:01:05.092Z"}
If InStr(TestResult, "clientOrderId") > 0 Then
    'Shouldn't happen with current test, for successfull orders
    Test.IsOk InStr(TestResult, "symbol") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsEqual JsonResult(1)("symbol"), "ETHBTC"
    Test.IsOk Len(JsonResult(1)("side")) >= 3
Else
    Test.IsOk InStr(TestResult, "message") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsEqual JsonResult("response_txt")("error")("code"), 2011
    Test.IsEqual JsonResult("response_txt")("error")("message"), "Quantity too low"
End If

End Sub

Function PublicHitBTCv2(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
Dim PayloadDict As New Dictionary

PublicApiSite = "https://api.hitbtc.com"

'Get special parameters currency and symbol and add them to the URL
If Not ParamDict Is Nothing Then
    For Each key In ParamDict.Keys
        If LCase(key) = "currency" Or LCase(key) = "symbol" Then
            Method = Method & "/" & ParamDict(key)
        Else
            PayloadDict(key) = ParamDict(key)
        End If
    Next key
End If


MethodParams = DictToString(PayloadDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/api/2/public/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicHitBTCv2 = WebRequestURL(url, ReqType)

End Function
Function PrivateHitBTCv2(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim url As String
Dim MethodParams As String

NonceUnique = CreateNonce(10)
TradeApiSite = "https://api.hitbtc.com"
urlPath = "/api/2/" & Method
MethodParams = DictToString(ParamDict, "URLENC")
postdata = JsonConverter.ConvertToJson(ParamDict)
If MethodParams <> "" Then MethodParams = "?" & MethodParams

url = TradeApiSite & urlPath

Dim headerDict As New Dictionary
headerDict.Add "Content-Type", "application/json"
'Credentials in a special format
headerDict.Add "Authorization", "Basic " & Base64Encode(Credentials("apiKey") & ":" & Credentials("secretKey"))

url = TradeApiSite & urlPath & MethodParams
PrivateHitBTCv2 = WebRequestURL(url, ReqType, headerDict, postdata)

End Function
