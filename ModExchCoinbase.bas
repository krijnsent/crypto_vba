Attribute VB_Name = "ModExchCoinbase"
Sub TestCoinbase()

'Standard Coinbase, for CoinbasePro (formerly known as GDAX), see that Module
'https://developers.coinbase.com/api/v2#introduction
'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 3 lines, unless you define 3 constants somewhere ( Public Const secretkey_gdax = "the key to use everywhere" etc )
apiKey = apikey_coinbase
secretKey = secretkey_coinbase

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchCoinbase"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestCoinbasePublic")

'Error, unknown command
TestResult = PublicCoinbase("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"errors":[{"id":"not_found","message":"Not found"}]}}
Test.IsOk InStr(TestResult, "error") > 0
Test.IsOk InStr(TestResult, "not_found") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Request wrong parameter
Dim Params As New Dictionary
Params.Add "currency", "X"
TestResult = PublicCoinbase("exchange-rates", "GET", Params)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"errors":[{"id":"invalid_request","message":"Invalid currency (X)"}]}}
Test.IsOk InStr(TestResult, "error") > 0
Test.IsOk InStr(TestResult, "invalid_request") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400

'Simpel request without parameters
TestResult = PublicCoinbase("currencies", "GET")
'{"data":[{"id":"AED","name":"United Arab Emirates Dirham","min_size":"0.01000000"},{"id":"AFN","name":"Afghan Afghani","min_size":"0.01000000"},{"id":"ALL","name":"Albanian Lek","min_size":"0.01000000"},
Test.IsOk InStr(TestResult, "min_size") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("data").Count >= 20
Test.IsEqual JsonResult("data")(1)("id"), "AED"
Test.IsEqual JsonResult("data")(1)("name"), "United Arab Emirates Dirham"
Test.IsEqual Val(JsonResult("data")(1)("min_size")), 0.01

'Request with parameter
Dim Params2 As New Dictionary
Params2.Add "currency", "ETH"
TestResult = PublicCoinbase("exchange-rates", "GET", Params2)
'{"data":{"currency":"ETH","rates":{"AED":"503.843775","AFN":"10260.72100155","ALL":"15205.84875","AMD":"66996.080561325","ANG":"250.3323036", etc
Test.IsOk InStr(TestResult, "EUR") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("data")("currency"), "ETH"
Test.IsEqual Val(JsonResult("data")("rates")("ETH")), 1
Test.IsOk Val(JsonResult("data")("rates")("USD")) > 0

'Coinbase time
TestResult = GetCoinbaseTime
Test.IsOk TestResult > 1550000000

Set Test = Suite.Test("TestCoinbasePrivate")
TestResult = PrivateCoinbase("accounts", "GET", Cred)
'{"pagination":{"ending_before":null,"starting_after":null,"limit":25,"order":"desc","previous_uri":null,"next_uri":null},"data":[{"id":"0cdbaac7-da83-5b85-0fe555be0b48","name":"EUR-wallet","primary":false,"type":"fiat","currency":{"code":"EUR","name":"Euro","color":"#0066cf","sort_index":0,"exponent":2,"type":"fiat"},"balance":{"amount":"0.00","currency":"EUR"},"created_at":"2017-12-27T16:57:41Z","updated_at":"2017-12-27T16:57:41Z","resource":"account","resource_path":"/v2/accounts/0cdbaac7-da83-5b85-b647-0fe402be0b48","allow_deposits":true,"allow_withdrawals":true},{"id":"0a3c2dfc-1c62-190b-abef-fbba3102c89b","name":"LTC-wallet","primary":true,"type":"wallet", etc...
Test.IsOk InStr(TestResult, "currency") > 0
Test.IsOk InStr(TestResult, "warnings") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("pagination")("limit"), 25
Test.IsOk JsonResult("data").Count >= 1
Test.IsEqual JsonResult("warnings")(1)("id"), "missing_version"

'user with CB-VERSION (API client version you can add to your requests to make sure you have the same version as you checked online, but no response is given
'Request with CB-VERSION
Dim Params3 As New Dictionary
Params3.Add "CB-VERSION", "2005-05-05"
TestResult = PrivateCoinbase("user", "GET", Cred, Params3)
'{"data":{"id":"3c7-12505bcbf174","name":"Koen Rijnsent","username":null,"profile_location":null,"profile_bio":null,"profile_url":null,"avatar_url":"https://res.cloudinary.com/coinbase/image/upload/c_fill,h_128,w_128/heg.png","resource":"user","resource_path":"/v2/user","email":"donotmailthis@here.com","time_zone":"Pacific Time (US \u0026 Canada)","native_currency":"EUR","bitcoin_unit":"BTC","state":null,"country":{"code":"NL","name":"Netherlands","is_in_europe":true},"region_supports_fiat_transfers":true,"region_supports_crypto_to_crypto_transfers":true,"created_at":"2008-01-01T16:51:09Z","tiers":{"completed_description":"Level 1","upgrade_button_text":null,"header":null,"body":null},"referral_money":{"amount":"8.90","currency":"EUR","currency_symbol":"€"}}}
Test.IsEqual InStr(TestResult, "warnings"), 0
Test.IsOk InStr(TestResult, "profile_location") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Len(JsonResult("data")("id")) > 10
Test.IsOk Len(JsonResult("data")("native_currency")) >= 3

'Update the default currency to EUR
Dim Params4 As New Dictionary
Params4.Add "CB-VERSION", "2005-05-05"
Params4.Add "native_currency", "EUR"
TestResult = PrivateCoinbase("user", "PUT", Cred, Params4)
'{"data":{"id":"3c7-12505bcbf174","name":"Koen Rijnsent","username":null,"profile_location":null,"profile_bio":null,"profile_url":null,"avatar_url":"https://res.cloudinary.com/coinbase/image/upload/c_fill,h_128,w_128/heg.png","resource":"user","resource_path":"/v2/user","email":"donotmailthis@here.com","time_zone":"Pacific Time (US \u0026 Canada)","native_currency":"EUR","bitcoin_unit":"BTC","state":null,"country":{"code":"NL","name":"Netherlands","is_in_europe":true},"region_supports_fiat_transfers":true,"region_supports_crypto_to_crypto_transfers":true,"created_at":"2008-01-01T16:51:09Z","tiers":{"completed_description":"Level 1","upgrade_button_text":null,"header":null,"body":null},"referral_money":{"amount":"8.90","currency":"EUR","currency_symbol":"€"}}}
Test.IsEqual InStr(TestResult, "warnings"), 0
Test.IsOk InStr(TestResult, "profile_location") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Len(JsonResult("data")("id")) > 10
Test.IsEqual JsonResult("data")("referral_money")("currency"), "EUR"

'Buy order that errors out
Dim Params5 As New Dictionary
Params5.Add "CB-VERSION", "2005-05-05"
Params5.Add "amount", 3
Params5.Add "currency", "BTC"
Params5.Add "quote", "true"
TestResult = PrivateCoinbase("accounts/the_right_account_here/buys", "POST", Cred, Params5)
'error with account: {"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"errors":[{"id":"invalid_request","message":"Can't buy with this account"}]}}
'unknown account id: {"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"errors":[{"id":"not_found","message":"Not found"}]}}
Test.IsOk InStr(TestResult, "errors") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("response_txt")("errors").Count >= 1


End Sub

Function PublicCoinbase(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://api.coinbase.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/v2/" & Method & MethodParams
Url = PublicApiSite & urlPath

PublicCoinbase = WebRequestURL(Url, ReqType)

End Function
Function PrivateCoinbase(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim Url As String
Dim CBVersion As String
Dim MethodParams As String

'Get a 10-digit Nonce
NonceUnique = GetCoinbaseTime
TradeApiSite = "https://api.coinbase.com/v2/"

'If a CB-VERSION is present, put it in a variable and remove it from the Parameter dictionary
CBVersion = ""
MethodParams = ""
If Not ParamDict Is Nothing Then
    If ParamDict.Exists("CB-VERSION") Then
        CBVersion = ParamDict("CB-VERSION")
        ParamDict.Remove "CB-VERSION"
    End If
    'Change the rest of the parameters to JSON
    MethodParams = DictToString(ParamDict, "JSON")
    If MethodParams = "{}" Then MethodParams = ""
End If

SignMsg = NonceUnique & UCase(ReqType) & "/v2/" & Method & MethodParams
APIsign = ComputeHash_C("SHA256", SignMsg, Credentials("secretKey"), "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/json"
headerDict.Add "CB-ACCESS-KEY", Credentials("apiKey")
headerDict.Add "CB-ACCESS-SIGN", APIsign
headerDict.Add "CB-ACCESS-TIMESTAMP", NonceUnique
If CBVersion <> "" Then
    headerDict.Add "CB-VERSION", CBVersion
End If

Url = TradeApiSite & Method
PrivateCoinbase = WebRequestURL(Url, ReqType, headerDict, MethodParams)


End Function

Function GetCoinbaseTime() As Double

Dim JsonResponse As String
Dim json As Object

JsonResponse = PublicCoinbase("time", "GET")
Set json = JsonConverter.ParseJson(JsonResponse)
GetCoinbaseTime = Int(json("data")("epoch"))
If GetCoinbaseTime = 0 Then
    TimeCorrection = -3600
    GetCoinbaseTime = CreateNonce(10)
    GetCoinbaseTime = Trim(Str((Val(GetCoinbaseTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set json = Nothing

End Function


