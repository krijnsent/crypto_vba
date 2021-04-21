Attribute VB_Name = "ModExchBitVavo"
Sub TestBitVavo()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://docs.bitvavo.com/
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_bitvavo
secretKey = secretkey_bitvavo

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBitVavo"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBitVavoPublic")

'Error, unknown command
TestResult = PublicBitVavo("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"errorCode":110,"error":"Invalid endpoint. Please check url and HTTP method."}}
Test.IsOk InStr(TestResult, "error") > 0, "test error 1a failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test error 1b failed, result: ${1}"

'Error, parameter missing
TestResult = PublicBitVavo("BTC-EUR/candles", "GET")
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"errorCode":203,"error":"interval parameter is required."}}
Test.IsOk InStr(TestResult, "error") > 0, "test error 2a failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400, "test error 2b failed, result: ${1}"
Test.IsEqual JsonResult("response_txt")("error"), "interval parameter is required.", "test error 2c failed, result: ${1}"

'OK simple time request
TestResult = PublicBitVavo("time", "GET")
'e.g. {"time":1617720826734}
Test.IsOk InStr(TestResult, "time") > 0, "test time 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("time") > 1600000000000#, "test time 2 failed, result: ${1}"

'OK request with parameter
Dim Params As New Dictionary
Params.Add "interval", "1d"
Params.Add "limit", 10
TestResult = PublicBitVavo("BTC-EUR/candles", "GET", Params)
'[[1617667200000,"49950","50300","48547","49010","455.91371112"],[1617580800000,"49590","50200","48500","49870","555.41905353"], etc.
'returns: time, OHLCV
Test.IsOk InStr(TestResult, "error") = 0, "test candles 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
For N = 1 To JsonResult.Count
    Test.IsOk JsonResult(N)(1) > 1600000000000#, "test candles 2-" & N & " failed, result: ${1}"   'check time of record
    Test.IsOk JsonResult(N)(2) > 0, "test candles 3-" & N & " failed, result: ${1}" 'check Open of record
Next N


Set Test = Suite.Test("TestBitVavoPrivate")
TestResult = PrivateBitVavo("account", "GET", Cred)
'e.g. {"fees":{"taker":"0.0025","maker":"0.0015","volume":"0.00"}}
Test.IsOk InStr(TestResult, "fees") > 0, "test private account 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("fees")("taker") >= 0, "test private account 2 failed, result: ${1}"

'Private GET request that requires a parameter
TestResult = PrivateBitVavo("deposit", "GET", Cred)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"errorCode":203,"error":"symbol parameter is required."}}
Test.IsOk InStr(TestResult, "error_txt") > 0, "test private deposit 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("response_txt")("error"), "symbol parameter is required.", "test private deposit 2 failed, result: ${1}"


Dim Params2 As New Dictionary
Params2.Add "symbol", "ETH"
TestResult = PrivateBitVavo("deposit", "GET", Cred, Params2)
'{"errorCode":412,"error":"crypto_bank_required."} - no deposit address set
'Or {"address": "CryptoCurrencyAddress","paymentId": "10002653"} - deposit address set
AddrSet = False
If InStr(TestResult, "address") > 0 Then AddrSet = True
Set JsonResult = JsonConverter.ParseJson(TestResult)

If AddrSet Then
    Test.IsOk JsonResult("address") <> "", "test private deposit 3b failed, result: ${1}"
Else
    Test.IsEqual JsonResult("response_txt")("error"), "crypto_bank_required.", "test private deposit 3a failed, result: ${1}"
End If

'Sign test case from API docs
TestMsgToSign = "1548172481125POST/v2/order{""market"":""BTC-EUR"",""side"":""buy"",""price"":""5000"",""amount"":""1.23"",""orderType"":""limit""}"
TestSign = ComputeHash_C("SHA256", TestMsgToSign, "bitvavo", "STRHEX")
Test.IsEqual TestSign, "44d022723a20973a18f7ee97398b9fdd405d2d019c8d39e24b8cc0dcb39ca016", "test sign failed, result: ${1}"


'Buy order, buying 1 BTC for 100 EUR/BTC
Dim Params3 As New Dictionary
Params3.Add "market", "BTC-EUR"
Params3.Add "side", "buy"
Params3.Add "orderType", "limit"
Params3.Add "amount", 1
Params3.Add "price", 100
Params3.Add "timeInForce", "FOK"
TestResult = PrivateBitVavo("order", "POST", Cred, Params3)
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"errorCode":216,"error":"You do not have sufficient balance to complete this operation."}}
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400, "test private order 1 failed, result: ${1}"
Test.IsEqual JsonResult("response_txt")("errorCode"), 216, "test private order 2 failed, result: ${1}"


'Deleting not existing order
Dim Params4 As New Dictionary
Params4.Add "market", "ETH-EUR"
Params4.Add "orderId", "ff403e21-e270-4584-bc9e-9c4b18461465"
TestResult = PrivateBitVavo("order", "DELETE", Cred, Params4)
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"errorCode":240,"error":"No order found. Please be aware that simultaneously updating the same order may return this error."}}
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test private delete order 1 failed, result: ${1}"
Test.IsEqual JsonResult("response_txt")("errorCode"), 240, "test private delete order 2 failed, result: ${1}"

'Test by default switched off... Deletes all open orders...
'Dim Params5 As New Dictionary
'TestResult = PrivateBitVavo("orders", "DELETE", Cred, Params5)
'TestResult = "{""orderId"": ""2e7ce7fc-44e2-4d80-a4a7-d079c4750b61""}"
'If InStr(TestResult, "orderId") > 0 Then
'    'has some results
'    'e.g.: {"orderId": "2e7ce7fc-44e2-4d80-a4a7-d079c4750b61"}
'    Test.IsOk InStr(TestResult, "orderId") > 0
'    Set JsonResult = JsonConverter.ParseJson(TestResult)
'    For Each k In JsonResult.Keys()
'        Test.IsOk Len(JsonResult(k)) >= 10
'    Next k
'Else
'    'no results
'    'Empty: []
'    Test.IsEqual TestResult, "[]"
'End If


End Sub

Function PublicBitVavo(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.bitvavo.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/v2/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBitVavo = WebRequestURL(url, ReqType)

End Function
Function PrivateBitVavo(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim MethodParams As String
Dim postdata As String
Dim url As String

TradeApiSite = "https://api.bitvavo.com"
ApiEndPoint = "/v2/" & Method
postdata = ""
NonceUnique = GetBitVavoTime

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    postdata = JsonConverter.ConvertToJson(ParamDict)
    If postdata = "{}" Then postdata = ""
ElseIf UCase(ReqType) = "GET" Or UCase(ReqType) = "DELETE" Or UCase(ReqType) = "PUT" Then
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    ApiEndPoint = ApiEndPoint & MethodParams
End If

StrToHash = NonceUnique & ReqType & ApiEndPoint & postdata
APIsign = ComputeHash_C("SHA256", StrToHash, Credentials("secretKey"), "STRHEX")
url = TradeApiSite & ApiEndPoint

Dim UrlHeaders As New Dictionary
UrlHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
UrlHeaders.Add "Content-Type", "application/json"
UrlHeaders.Add "BITVAVO-ACCESS-TIMESTAMP", NonceUnique
UrlHeaders.Add "BITVAVO-ACCESS-KEY", Credentials("apiKey")
UrlHeaders.Add "BITVAVO-ACCESS-SIGNATURE", APIsign
PrivateBitVavo = WebRequestURL(url, ReqType, UrlHeaders, postdata)

End Function


Function GetBitVavoTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicBinance time
JsonResponse = PublicBitVavo("time", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetBitVavoTime = Json("time")

Set Json = Nothing

End Function


