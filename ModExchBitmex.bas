Attribute VB_Name = "ModExchBitmex"
Sub TestBitmex()

'Source: https://github.com/krijnsent/crypto_vba
'Documentation: https://www.bitmex.com/app/restAPI
'Commands: https://www.bitmex.com/api/explorer/
'VBA example: https://github.com/BitMEX/api-connectors/tree/master/official-http/vba
'Remember to create a new API key for excel/VBA

Dim Apikey As String
Dim secretKey As String

Apikey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
Apikey = apikey_bitmex
secretKey = secretkey_bitmex

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", Apikey
Cred.Add "secretKey", secretKey

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModExchBitmex"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestBitmexPublic")

'Error, unknown command
TestResult = PublicBitmex("AnUnknownCommand", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":{"message":"Not Found","name":"HTTPError"}}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, command without parameters
TestResult = PublicBitmex("orderBook/L2", "GET")
'{"error_nr":400,"error_txt":"HTTP-Bad Request","response_txt":{"error":{"message":"'symbol' is a required arg.","name":"HTTPError"}}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400

'OK request
TestResult = PublicBitmex("stats", "GET")
'[{"rootSymbol":"A50","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"ADA","currency":"XBt","volume24h":28782927,"turnover24h":17393857814,"openInterest":54769214,"openValue":33902143466},{"rootSymbol":"BCH","currency":"XBt","volume24h":3642,"turnover24h":9362243000,"openInterest":24992,"openValue":64404384000},{"rootSymbol":"BFX","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"BLOCKS","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"BVOL","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"COIN","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"DAO","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0},{"rootSymbol":"DASH","currency":"XBt","volume24h":0,"turnover24h":0,"openInterest":0,"openValue":0} etc.
Test.IsOk InStr(TestResult, "ETH") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
For N = 1 To JsonResult.Count
    Test.IsEqual JsonResult(N)("currency"), "XBt"
    If JsonResult(N)("rootSymbol") <> "Total" Then Test.IsOk JsonResult(N)("volume24h") >= 0
Next N

'Put parameters/options in a dictionary
Dim Params As New Dictionary
Params.Add "symbol", "XBT"
Params.Add "depth", 5
TestResult = PublicBitmex("orderBook/L2", "GET", Params)
'[{"symbol":"XBTUSD","id":8799115700,"side":"Sell","size":65300,"price":8843},{"symbol":"XBTUSD","id":8799115750,"side":"Sell","size":58655,"price":8842.5},{"symbol":"XBTUSD","id":8799115800,"side":"Sell","size":88599,"price":8842},{"symbol":"XBTUSD","id":8799115850,"side":"Sell","size":5368,"price":8841.5},{"symbol":"XBTUSD","id":8799115900,"side":"Sell","size":1436605,"price":8841},{"symbol":"XBTUSD","id":8799115950,"side":"Buy","size":2230982,"price":8840.5},{"symbol":"XBTUSD","id":8799116000,"side":"Buy","size":30155,"price":8840},{"symbol":"XBTUSD","id":8799116050,"side":"Buy","size":61062,"price":8839.5},{"symbol":"XBTUSD","id":8799116100,"side":"Buy","size":78279,"price":8839},{"symbol":"XBTUSD","id":8799116150,"side":"Buy","size":81493,"price":8838.5}]
Test.IsOk InStr(TestResult, "symbol") > 0
Test.IsOk InStr(TestResult, "side") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult(1)("symbol"), "XBTUSD"
Test.IsOk JsonResult(1)("id") > 0
Test.IsOk JsonResult(1)("size") > 0
Test.IsOk JsonResult(1)("price") > 0

'GET private API
Set Test = Suite.Test("TestBitmexPrivate GET")

'Use TESTNET

'Test an invalid command
Dim Params2 As New Dictionary
Params2.Add "testnet", 1
TestResult = PrivateBitmex("not_a_command", "GET", Cred, Params2)
'{"error_nr":404,"error_txt":"HTTP-Not Found","response_txt":{"error":{"message":"Not Found","name":"HTTPError"}}}
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Simple GET without parameters
Dim Params3 As New Dictionary
Params3.Add "testnet", 1
TestResult = PrivateBitmex("user", "GET", Cred, Params3)
'{"id":30219,"ownerId":null,"lastname":"Rijnsent","username":"rijnsent","email":"rijnsent",etc..}
Test.IsOk InStr(TestResult, "lastname") > 0
Test.IsOk InStr(TestResult, "username") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("id") > 0

'Simple GET without parameters
Dim Params4 As New Dictionary
Params4.Add "testnet", 1
Params4.Add "currency", "XBt"
Params4.Add "count", 5
TestResult = PrivateBitmex("user/walletHistory", "GET", Cred, Params4)
'[{"transactID":"db7925ad-b54156-baff28-baf7","account":3210,"currency":"XBt","transactType":"Transfer","amount":1000000,"fee":null,"transactStatus":"Completed","address":"0","tx":"9ddad751-507a-81ca-0b55-13cd08b7063f","text":"Signup bonus","transactTime":"2020-06-01T18:14:33.791Z","walletBalance":1000000,"marginBalance":null,"timestamp":"2020-06-01T18:14:33.791Z"}]
Test.IsOk InStr(TestResult, "transactID") > 0
Test.IsOk InStr(TestResult, "currency") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult(1)("amount") > 0


Set Test = Suite.Test("TestBitmexPrivate POST/DELETE")
'Test delete all orders
Dim Params5 As New Dictionary
Params5.Add "testnet", 1
TestResult = PrivateBitmex("order/all", "DELETE", Cred, Params5)
Test.IsEqual TestResult, "[]"

'Test delete all orders
Dim Params6 As New Dictionary
Params6.Add "testnet", 1
Params6.Add "symbol", "XBTUSD"
Params6.Add "price", 2
Params6.Add "orderQty", 0.00000002
Params6.Add "clOrdID", "MyTestOrderIDHere"
TestResult = PrivateBitmex("order", "POST", Cred, Params6)
'{"error_nr":403,"error_txt":"HTTP-Forbidden","response_txt":{"error":{"message":"Access Denied","name":"HTTPError"}}}
Test.IsOk InStr(TestResult, "error_nr") > 0
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 403


End Sub

Function PublicBitmex(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://www.bitmex.com"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/api/v1/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicBitmex = WebRequestURL(url, ReqType)

End Function
Function PrivateBitmex(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim MethodParams As String
Dim postdata As String
Dim url As String

TradeApiSite = "https://www.bitmex.com"
If Not ParamDict Is Nothing Then
    If ParamDict.Exists("testnet") Then
        ParamDict.Remove "testnet"
        TradeApiSite = "https://testnet.bitmex.com"
    End If
End If
ApiEndPoint = "/api/v1/" & Method
postdata = ""
NonceUnique = CreateNonce(13)

If UCase(ReqType) = "POST" Then
    'For POST request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}).
    postdata = JsonConverter.ConvertToJson(ParamDict)
ElseIf UCase(ReqType) = "GET" Then
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
    ApiEndPoint = ApiEndPoint & MethodParams
End If


StrToHash = ReqType & ApiEndPoint & NonceUnique & postdata
APIsign = ComputeHash_C("SHA256", StrToHash, Credentials("secretKey"), "STRHEX")
url = TradeApiSite & ApiEndPoint

Dim UrlHeaders As New Dictionary
UrlHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
UrlHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
UrlHeaders.Add "api-nonce", NonceUnique 'NOT USED ANYMORE
UrlHeaders.Add "api-key", Credentials("apiKey")
UrlHeaders.Add "api-signature", APIsign
PrivateBitmex = WebRequestURL(url, ReqType, UrlHeaders, postdata)

End Function
