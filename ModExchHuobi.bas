Attribute VB_Name = "ModExchHuobi"
Sub TestHuobi()

'Source: https://github.com/krijnsent/crypto_vba
'https://alphaex-api.github.io/openapi/spot/v1/en/#introduction
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_huobi
secretKey = secretkey_huobi

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchHuobi"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestHuobiPublic")

'Error, unknown command
TestResult = PublicHuobi("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_txt"), "HTTP-Not Found"
Test.IsEqual JsonResult("error_nr"), 404

'OK request
TestResult = PublicHuobi("v1/common/timestamp", "GET")
'e.g. {"status":"ok","data":1579706923783}
Test.IsOk InStr(TestResult, "data") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "ok"
Test.IsOk Val(JsonResult("data")) > 1500000000000#

'Parameters missing
TestResult = PublicHuobi("market/history/kline", "GET")
'e.g. {"ts":1579707152954,"status":"error","err-code":"invalid-parameter","err-msg":"invalid symbol"}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "error"
Test.IsEqual JsonResult("err-msg"), "invalid symbol"
Test.IsOk Val(JsonResult("ts")) > 1500000000000#

'Put parameters/options in a dictionary
'If no parameters are provided, the defaults are used
Dim Params As New Dictionary
Params.Add "period", "1day"
Params.Add "symbol", "btcusdt"
Params.Add "size", 10
TestResult = PublicHuobi("market/history/kline", "GET", Params)
'e.g. {"status":"ok","ch":"market.btcusdt.kline.1day","ts":1579707654120,"data":[{"amount":25326.647313510339831018,"open":8645.130000000000000000,"close":8659.620000000000000000,"high":8817.730000000000000000,"id":1579622400,"count":202979,"low":8500.000000000000000000,"vol":219864523.567705105282063018560000000000000000},{"amount":17344.079067910875891838,"open":8677.970000000000000000,"close":8646.800000000000000000,"high":8744.510000000000000000,"id":1579536000,"count":153447,"low":8607.430000000000000000,"vol":150214939.669388488950943200910000000000000000},{"amount":27195.320357908427801956,"open":8632.820000000000000000,"close":8677.200000000000000000,"high":8756.040000000000000000,"id":1579449600,"count":234172,"low":8480.000000000000000000,"vol":235036539.681727868489706633830000000000000000},
Test.IsOk InStr(TestResult, "market.btcusdt.kline.1day") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk Val(JsonResult("ts")) > 1500000000000#
Test.IsEqual JsonResult("status"), "ok"
Test.IsEqual JsonResult("data").Count, 10
Test.IsOk Val(JsonResult("data")(1)("amount")) > 0
Test.IsOk Val(JsonResult("data")(2)("high")) > 0


Set Test = Suite.Test("TestHuobiPrivate GET")
'Simple test, should return data
TestResult = PrivateHuobi("v1/account/accounts", "GET", Cred)
'{"status":"ok","data":[{"id":9999,"type":"spot","subtype":"","state":"working"}]}
Debug.Print TestResult
Test.IsOk InStr(TestResult, "status") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "ok"
Test.IsEqual JsonResult("data")(1)("state"), "working"

'Error, forgotten parameter
Dim Params2 As New Dictionary
Params2.Add "size", 10
TestResult = PrivateHuobi("v1/account/history", "GET", Cred, Params2)
'{"status":"error","err-code":"validation-constraints-required","err-msg":"Field is missing: account-id.","data":null}
Test.IsOk InStr(TestResult, "err-msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "error"
Test.IsEqual JsonResult("err-msg"), "Field is missing: account-id."

'Unknown account-id
Dim Params3 As New Dictionary
Params3.Add "account-id", 9999
Params3.Add "size", 50
TestResult = PrivateHuobi("v1/account/history", "GET", Cred, Params3)
'{"status":"error","err-code":"account-get-balance-account-inexistent-error","err-msg":"account for id `6,000,006` and user id `9,999` does not exist","data":null}
Test.IsOk InStr(TestResult, "err-msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "error"
Test.IsEqual JsonResult("err-code"), "account-get-balance-account-inexistent-error"


Set Test = Suite.Test("TestHuobiPrivate POST")
'Get account-id:
TestResult = PrivateHuobi("v1/account/accounts", "GET", Cred)
Set JsonResult = JsonConverter.ParseJson(TestResult)
    AccId = JsonResult("data")(1)("id")
'Place order
Dim Params4 As New Dictionary
Params4.Add "account-id", AccId
Params4.Add "amount", 1
Params4.Add "price", 1
Params4.Add "symbol", "ethusdt"
Params4.Add "type", "buy-limit"
TestResult = PrivateHuobi("v1/order/orders/place", "POST", Cred, Params4)
'{"status":"error","err-code":"order-value-min-error","err-msg":"Order total cannot be lower than: `5`","data":null}
Test.IsOk InStr(TestResult, "status") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "error"
Test.IsEqual JsonResult("err-code"), "order-value-min-error"



End Sub

Function PublicHuobi(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api-cloud.huobi.co.kr/"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = Method & MethodParams
url = PublicApiSite & urlPath

'Debug.Print Url

PublicHuobi = WebRequestURL(url, ReqType)

End Function
Function PrivateHuobi(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim APIsign As String
Dim ApiEndPoint As String
Dim postdata As String
Dim url As String

'Get a Timestamp
Stamp = GetUTCTime()
StampTxt = URLEncode(Format(Stamp, "YYYY-MM-DDThh:mm:ss"))

HostTxt = "api-cloud.huobi.co.kr"
HostTxt = "api.huobi.pro"
TradeApiSite = "https://" & HostTxt & "/"

url = TradeApiSite & Method
StrHash = ""
postdata = ""

Dim TotDict As New Dictionary
TotDict.Add "AccessKeyId", Credentials("apiKey")
TotDict.Add "SignatureMethod", "HmacSHA256"
TotDict.Add "SignatureVersion", 2
TotDict.Add "Timestamp", StampTxt

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    MethodParams = DictToString(TotDict, "URLENC")
    
    postdata = JsonConverter.ConvertToJson(ParamDict)
    'ApiEndPoint = Url
ElseIf UCase(ReqType) = "GET" Then
    If Not ParamDict Is Nothing Then
        For Each key In ParamDict.Keys
            TotDict(key) = ParamDict(key)
        Next key
    End If
    MethodParams = DictToString(TotDict, "URLENC")
    postdata = ""
End If

StrHash = UCase(ReqType) & Chr(10) & HostTxt & Chr(10) & "/" & Method & Chr(10) & MethodParams

If MethodParams <> "" Then MethodParams = "?" & MethodParams
ApiEndPoint = url & MethodParams


'Dim PostDict As New Dictionary
'PostDict.Add "access_token", Credentials("apiKey")
'PostDict.Add "nonce", NonceUnique
'If Not ParamDict Is Nothing Then
'    For Each Key In ParamDict.Keys
'        PostDict(Key) = ParamDict(Key)
'    Next Key
'End If
'postdataUrl = DictToString(PostDict, "URLENC")
'postdataJSON = JsonConverter.ConvertToJson(ParamDict)
'postdata64 = Base64Encode(postdataJSON)

APIsign = ComputeHash_C("SHA256", StrHash, Credentials("secretKey"), "STR64")
APIsignEnc = URLEncode(APIsign)
ApiEndPoint = ApiEndPoint & "&Signature=" & APIsignEnc

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/json"

'Debug.Print ApiEndPoint

PrivateHuobi = WebRequestURL(ApiEndPoint, ReqType, headerDict, postdata)

End Function
