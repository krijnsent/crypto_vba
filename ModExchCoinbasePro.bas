Attribute VB_Name = "ModExchCoinbasePro"
Sub TestCoinbasePro()

'CoinbasePro, formerly known as GDAX
'For normal Coinbase, see the Coinbase API
'API docs: https://docs.pro.coinbase.com/
'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String
Dim passphrase As String

apiKey = "your api key here"
secretKey = "your secret key here"
passphrase = "your passphrase here"

'Remove these 3 lines, unless you define 3 constants somewhere ( Public Const secretkey_gdax = "the key to use everywhere" etc )
apiKey = apikey_coinbase_pro
secretKey = secretkey_coinbase_pro
passphrase = passphrase_coinbase_pro

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey
Cred.Add "Passphrase", passphrase

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchCoinbasePro"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestCoinbaseProPublic")

'Error, unknown command
TestResult = PublicCoinbasePro("AnUnknownCommand", "GET")
'{"error_nr":401,"error_txt":"HTTP-Unauthorized","response_txt":{"message":"CB-ACCESS-KEY header is required"}}
Test.IsOk InStr(TestResult, "message") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 401
Test.IsEqual JsonResult("response_txt")("message"), "CB-ACCESS-KEY header is required"

'Request wrong parameters
Dim Params As New Dictionary
Params.Add "level", 5
TestResult = PublicCoinbasePro("products/BTC-USD/book", "GET", Params)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"message":"Bad Request"}}
Test.IsOk InStr(TestResult, "message") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400
Test.IsEqual JsonResult("response_txt")("message"), "Bad Request"

'Request with parameter
Dim Params2 As New Dictionary
Params2.Add "level", 1
TestResult = PublicCoinbasePro("products/ETH-EUR/book", "GET", Params2)
'{"sequence":2052119022,"bids":[["118.04","200.16128756",5]],"asks":[["118.05","30.06104554",4]]}
Test.IsOk InStr(TestResult, "asks") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("sequence") > 1
Test.IsEqual JsonResult("bids").Count, 1
Test.IsEqual JsonResult("asks").Count, 1

'Coinbase time
TestResult = GetCoinbaseProTime
Test.IsOk TestResult > 1550000000

Set Test = Suite.Test("TestCoinbaseProPrivate")
TestResult = PrivateCoinbasePro("accounts", "GET", Cred)
'[{"id":"8a06fcff-f233-4b2a-b333-ec2ccd727956","currency":"BTC","balance":"0.0000000000000000","available":"0","hold":"0.0000000000000000","profile_id":"2c-015-61806709e17"},{"id":"b9d028fa-748a-9fa3-9df9b877457d","currency":"LTC","balance":"0.0000000000000000","available":"0","hold":" etc...
Test.IsOk InStr(TestResult, "profile_id") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult.Count > 1
Test.IsEqual JsonResult(1)("currency"), "BAT"
Test.IsOk JsonResult(1)("balance") >= 0

Dim Params8 As New Dictionary
Params8.Add "size", 0.01
Params8.Add "price", 100.1
Params8.Add "side", "buy"
Params8.Add "product_id", "BTC-EUR"
TestResult = PrivateCoinbasePro("orders", "POST", Cred, Params8)
If InStr(TestResult, "error_txt") > 0 Then
    'Error result, assume insufficient funds, but could also be Product not found
    '{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"message":"Insufficient funds"}}
    Test.IsOk InStr(TestResult, "response_txt") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsEqual JsonResult("response_txt")("message"), "Insufficient funds"
Else
    'Normal result
    '{"id": "d0c5340b-6d6c-49d9-b567-48c4bfca13d2","price": "100.10000000","size": "0.01000000","product_id": "BTC-EUR","side": "buy","stp": "dc","type": "limit","time_in_force": "GTC","post_only": false,"created_at": "2016-12-08T20:02:28.53864Z","fill_fees": "0.0000000000000000","filled_size": "0.00000000","executed_value": "0.0000000000000000","status": "pending","settled": false}
    Test.IsOk InStr(TestResult, "created_at") > 0
    Set JsonResult = JsonConverter.ParseJson(TestResult)
    Test.IsOk Len(JsonResult("id")) > 10
    Test.IsEqual JsonResult("product_id"), "BTC-EUR"
End If


'Delete all BTC-EUR orders
Dim Params3 As New Dictionary
Params3.Add "product_id", "BTC-EUR"
TestResult = PrivateCoinbasePro("orders", "DELETE", Cred, Params3)
'No orders to delete: []
Test.IsEqual TestResult, "[]"


'Withdraw one BAT to an invalid account
Dim Params4 As New Dictionary
Params4.Add "amount", 1
Params4.Add "currency", "BAT"
Params4.Add "crypto_address", "0x0"
TestResult = PrivateCoinbasePro("withdrawals/crypto", "POST", Cred, Params4)
'E.g. {"error_nr":403,"error_txt":"HTTP-Forbidden","response_txt":{"message":"Forbidden"}}
Test.IsOk InStr(TestResult, "Forbidden") > 0


End Sub

Function PublicCoinbasePro(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.pro.coinbase.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicCoinbasePro = WebRequestURL(url, ReqType)

End Function
Function PrivateCoinbasePro(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim url As String
Dim MethodParams As String

'Get a 10-digit Nonce
NonceUnique = GetCoinbaseProTime
TradeApiSite = "https://api.pro.coinbase.com"

'Change the parameters to JSON
MethodParams = DictToString(ParamDict, "JSON")
If MethodParams = "{}" Then MethodParams = ""
    
SignMsg = NonceUnique & UCase(ReqType) & "/" & Method & "" & MethodParams
APIsign = Base64Encode(ComputeHash_C("SHA256", SignMsg, Base64Decode(Credentials("secretKey")), "RAW"))

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/json"
headerDict.Add "CB-ACCESS-KEY", Credentials("apiKey")
headerDict.Add "CB-ACCESS-SIGN", APIsign
headerDict.Add "CB-ACCESS-TIMESTAMP", NonceUnique
headerDict.Add "CB-ACCESS-PASSPHRASE", Credentials("Passphrase")

url = TradeApiSite & "/" & Method
PrivateCoinbasePro = WebRequestURL(url, ReqType, headerDict, MethodParams)

End Function

Function GetCoinbaseProTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicCoinbasePro time
JsonResponse = PublicCoinbasePro("time", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetCoinbaseProTime = Int(Json("epoch"))
If GetCoinbaseProTime = 0 Then
    TimeCorrection = -3600
    GetCoinbaseProTime = CreateNonce(10)
    GetCoinbaseProTime = Trim(Str((Val(GetGDAXTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set Json = Nothing

End Function

