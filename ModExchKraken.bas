Attribute VB_Name = "ModExchKraken"
Sub TestKraken()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Kraken will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources
'https://www.kraken.com/en-us/help/api#public-market-data
'https://www.kraken.com/help/api#private-user-data

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_kraken = "the key to use everywhere" etc )
apiKey = apikey_kraken
secretKey = secretkey_kraken

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchKraken"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestKrakenPublic")

'Error, unknown command
TestResult = PublicKraken("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":["EGeneral:Unknown method"]}}
Test.IsOk InStr(TestResult, "error") > 0, "test error 1 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404, "test error 2 failed, result: ${1}"

'Error, parameter missing
TestResult = PublicKraken("Ticker", "GET")
'{"error":["EGeneral:Invalid arguments"]}
Test.IsOk InStr(TestResult, "Invalid") > 0, "test error 3 failed, result: ${1}"
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error")(1), "EGeneral:Invalid arguments", "test error 4 failed, result: ${1}"

'Ok request without parameters
TestResult = PublicKraken("Time", "GET")
'Example: {"error":[],"result":{"unixtime":1551737935,"rfc1123":"Mon,  4 Mar 19 22:18:55 +0000"}}
Test.IsOk InStr(TestResult, "unixtime") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result")("unixtime") >= 1510000000

Dim Params As New Dictionary
Params.Add "pair", "XXBTZEUR"
TestResult = PublicKraken("OHLC", "GET", Params)
'{"error":[],"result":{"XXBTZEUR":[[1551695100,"3265.8","3265.8","3265.2","3265.2","3265.5","0.53688049",12],[1551695160,"3265.2", etc...
Test.IsOk InStr(TestResult, "XXBTZEUR") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result")("XXBTZEUR")(1)(1) >= 1510000000

Set Test = Suite.Test("TestKrakenPrivate")
TestResult = PrivateKraken("Balance", "POST", Cred)
'{"error":[],"result":{"ZEUR":"15.35","KFEE":"935","XXBT": etc...
Test.IsOk InStr(TestResult, "ZEUR") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result")("KFEE") >= 0

'Unix time period:
t1 = DateToUnixTime("1/1/2016")
t2 = DateToUnixTime("1/1/2018")

Dim Params2 As New Dictionary
Params2.Add "start", t1
Params2.Add "end", t2
TestResult = PrivateKraken("TradesHistory", "POST", Cred, Params2)
'{"error":[],"result":{"trades":{"TBSI6I-EO4KN-MLU4AI":{"ordertxid":"O7AERY-NCNDR-6WKLMU","pair":"XXMRZEUR","time":1493715960.4854,"type":"buy","ordertype":"limit","price": etc...
Test.IsOk InStr(TestResult, "trades") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("result")("trades").Count >= 0


End Sub

Function PublicKraken(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.kraken.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/0/public/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicKraken = WebRequestURL(url, ReqType)

End Function
Function PrivateKraken(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim postdata As String
Dim PayloadDict As Dictionary
Dim url As String

'Kraken nonce: 16 characters
NonceUnique = CreateNonce(16)

TradeApiSite = "https://api.kraken.com"
urlPath = "/0/private/" & Method

Set PayloadDict = New Dictionary
If Not ParamDict Is Nothing Then
    For Each key In ParamDict.Keys
        PayloadDict(key) = ParamDict(key)
    Next key
End If
PayloadDict("nonce") = NonceUnique
postdata = DictToString(PayloadDict, "URLENC")

url = TradeApiSite & urlPath
APIsign = ComputeHash_C("SHA512", urlPath & ComputeHash_C("SHA256", NonceUnique & postdata, "", "RAW"), Base64Decode(Credentials("secretKey")), "STR64")

Dim headerDict As New Dictionary
headerDict.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "API-Key", Credentials("apiKey")
headerDict.Add "API-Sign", APIsign

PrivateKraken = WebRequestURL(url, ReqType, headerDict, postdata)

End Function
