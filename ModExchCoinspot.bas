Attribute VB_Name = "ModExchCoinspot"
Sub TestCoinspot()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://www.coinspot.com.au/api
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_coinspot
secretKey = secretkey_coinspot

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchCoinspot"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestCoinspotPublic")

'Error, unknown command
TestResult = PublicCoinspot("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":0}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Request without parameters (for Coinspot only public request)
TestResult = PublicCoinspot("latest", "GET")
'{"status":"ok","prices":{"btc":{"bid":"5330.10000001","ask":"5394","last":"5367"},"ltc":{"bid":"67.1","ask":"68.7","last":"68"},"doge":{"bid":"0.0027","ask":"0.0028","last":"0.0028"},"eth":{"bid":"186.11","ask":"191.99","last":"187"},"powr":{"bid":"0.133","ask":"0.1425","last":"0.14"},"ans":{"bid":"12.5","ask":"13","last":"12.5"},"xrp":{"bid":"0.44","ask":"0.449","last":"0.442"},"trx":{"bid":"0.0325","ask":"0.033999","last":"0.0327"}}}
Test.IsOk InStr(TestResult, "btc") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "ok"
Test.IsOk JsonResult("prices").Count >= 3
Test.IsOk JsonResult("prices")("btc")("last") > 0


Set Test = Suite.Test("TestCoinspotPrivate")
TestResult = PrivateCoinspot("my/balances", "POST", Cred)
'e.g. {"status":"ok","balance":{"btc":0,"ltc":3,"doge":1000,"ppc":0,"wdc":0,"xpm":0,"max":0,"lot":0,"qrk":0,"moon":0,"ftc":0,"drk":0}}
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "ok"
Test.IsOk JsonResult("balance")("btc") >= 0
Test.IsOk JsonResult("balance")("doge") >= 0

'Put the parameters in a dictionary
Dim Params As New Dictionary
Params.Add "cointype", "DOGE"
Params.Add "amount", 10000
TestResult = PrivateCoinspot("quote/buy", "POST", Cred, Params)
'e.g. {"status":"ok","quote":0.001619,"timeframe":0}
Test.IsOk InStr(TestResult, "quote") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("status"), "ok"
Test.IsOk JsonResult("quote") >= 0
Test.IsOk JsonResult("timeframe") >= 0


End Sub

Function PublicCoinspot(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim Url As String
PublicApiSite = "https://www.coinspot.com.au"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "&" & MethodParams
urlPath = "/pubapi/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicCoinspot = WebRequestURL(Url, ReqType)

End Function
Function PrivateCoinspot(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim Url As String
Dim PayloadDict As Dictionary
Dim MethodParams As String

'Get a Nonce
NonceUnique = CreateNonce()
TradeApiSite = "https://www.coinspot.com.au"

Set PayloadDict = New Dictionary
PayloadDict("nonce") = val(NonceUnique)
If Not ParamDict Is Nothing Then
    For Each Key In ParamDict.Keys
        PayloadDict(Key) = ParamDict(Key)
    Next Key
End If
MethodParams = DictToString(PayloadDict, "JSON")

PostPath = "/api/" & Method
APIsign = ComputeHash_C("SHA512", MethodParams, Credentials("secretKey"), "STRHEX")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/json"
headerDict.Add "sign", APIsign
headerDict.Add "key", Credentials("apiKey")

Url = TradeApiSite & PostPath
PrivateCoinspot = WebRequestURL(Url, "POST", headerDict, MethodParams)

End Function


