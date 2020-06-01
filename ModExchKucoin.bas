Attribute VB_Name = "ModExchKucoin"
Sub TestKucoin()

'Source: https://github.com/krijnsent/crypto_vba
'https://docs.kucoin.com/
'Remember to create a new API key for excel/VBA
'Kucoin will require ever increasing values/nonces for the private API and the nonces created in VBA might mismatch that of other sources

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_Kucoin = "the key to use everywhere" etc )
apiKey = apikey_kucoin
secretKey = secretkey_kucoin
passphrase = passphrase_kucoin

'Put the credentials in a dictionary
Dim Cred As New Dictionary
Cred.Add "apiKey", apiKey
Cred.Add "secretKey", secretKey
Cred.Add "Passphrase", passphrase

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
TestResult = PublicKucoin("AnUnknownCommand", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 404

'Error, missing parameters
TestResult = PublicKucoin("market/orderbook/level1", "GET")
Test.IsOk InStr(TestResult, "error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("error_nr"), 400

TestResult = PublicKucoin("market/allTickers", "GET")
'{"code":"200000","data":{"ticker":[{"symbol":"LOOM-BTC","high":"0.00001204","vol":"39738.31683935","last":"0.00001187","low":"0.00001151","buy":"0.00001172","sell":"0.00001187","changePrice":"0.00000025","changeRate":"0.0215"},etc...
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "changePrice") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data")("ticker").Count > 100
Test.IsOk Len(JsonResult("data")("ticker")(9)("symbol")) > 0
Test.IsOk JsonResult("data")("ticker")(3)("vol") > 0

Dim Params As New Dictionary
Params.Add "symbol", "KCS-BTC"
TestResult = PublicKucoin("market/orderbook/level2_20", "GET", Params)
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

TestResult = PrivateKucoin("accounts", "GET", Cred)
'{"code":"200000","data":[{"balance":"15.827819","available":"15.827819","holds":"0","currency":"KCS","id":"5c6a4a1d81a34e1da97","type":"trade"},{"balance":"2.12058951","available":"2.12058951",", etc...
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data").Count > 20
Test.IsOk JsonResult("data")(1)("balance") > 0

'Get only KCS account amount
Dim Params1 As New Dictionary
Params1.Add "currency", "KCS"
TestResult = PrivateKucoin("accounts", "GET", Cred, Params1)
'Debug.Print TestResult
'{"code":"200000","data":[{"balance":"15.82887819","available":"15.82887819","holds":"0","currency":"KCS","id":"5c6a4a1d81a34e1da97","type":"trade"}]}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "balance") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 200000
Test.IsOk JsonResult("data").Count >= 1
Test.IsOk JsonResult("data")(1)("balance") > 0

'Create a main LTC account (if it doesn't exist)
Dim Params2 As New Dictionary
Params2.Add "currency", "LTC"
TestResult = PrivateKucoin("accounts", "POST", Cred, Params2)
'Debug.Print TestResult
'{"code":"400100","msg":"type can not be empty"}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "msg") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("code") * 1, 400100
Test.IsEqual JsonResult("msg"), "type can not be empty"

Params2.Add "type", "main"
TestResult = PrivateKucoin("accounts", "POST", Cred, Params2)
'Debug.Print TestResult
'FIRST TIME RESULT: {"code":"200000","data":{"id":"5c7556e3cbfc7b24a1a1a1a9"}}
'NEXT RESULT: {"code":"230005","msg":"account already exists"}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "msg") + InStr(TestResult, "data") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 230005

Set Test = Suite.Test("TestKucoinPrivate Orders")
'Create orders
'sell 0.01 KCS for a price of 100 KCS per ETH
'price hopefully insane enough never to execute
TempOrderID = CreateNonce()
Dim Params3 As New Dictionary
Params3.Add "clientOid", TempOrderID
Params3.Add "symbol", "KCS-ETH"
Params3.Add "side", "sell"
Params3.Add "price", 100
Params3.Add "size", 0.01
Params3.Add "timeInForce", "GTC"
TestResult = PrivateKucoin("orders", "POST", Cred, Params3)
'Debug.Print TestResult
'{"code":"200000","data":{"orderId":"5ca22ec6513ab9576fb77d92"}}
'{"code":"200004","msg":"Balance insufficient!"}
Test.IsOk InStr(TestResult, "code") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004

'Add another order
Dim Params4 As New Dictionary
Params4.Add "clientOid", TempOrderID + 3
Params4.Add "symbol", "KCS-BTC"
Params4.Add "side", "sell"
Params4.Add "price", 100
Params4.Add "size", 0.01
Params4.Add "timeInForce", "GTC"
TestResult = PrivateKucoin("orders", "POST", Cred, Params4)
'Debug.Print TestResult
'{"code":"200000","data":{"orderId":"5ca22ec6513ab9576fb77d92"}}
'{"code":"200004","msg":"Balance insufficient!"}
Test.IsOk InStr(TestResult, "code") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004

'Now get the open orders
TestResult = PrivateKucoin("orders", "GET", Cred)
'{"code":"200000","data":{"totalNum":8,"totalPage":1,"pageSize":50,"currentPage":1,"items":[{"symbol":"KCS-BTC","hidden":false,"opType":"DEAL"
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "totalPage") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004
Test.IsOk JsonResult("data")("items").Count >= 0

'Delete all KCS-BTC orders
Dim Params5 As New Dictionary
Params5.Add "symbol", "KCS-BTC"
TestResult = PrivateKucoin("orders", "DELETE", Cred, Params5)
'{"code":"200000","data":{"cancelledOrderIds":["5ca2798389fc8450590fe207"]}}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "cancelledOrderIds") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004

'Delete the created KCS-ETH order
Dim Params6 As New Dictionary
Params6.Add "OrderId", JsonResult("data")("orderId")
TestResult = PrivateKucoin("orders", "DELETE", Cred, Params6)
'Debug.Print TestResult
'{"code":"200000","data":{"cancelledOrderIds":["5ca27982054b467eb0d0c8dc"]}}
Test.IsOk InStr(TestResult, "code") > 0
Test.IsOk InStr(TestResult, "cancelledOrderIds") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004

'Delete all orders (should be none)
TestResult = PrivateKucoin("orders", "DELETE", Cred)
'{"code":"200000","data":{"cancelledOrderIds":[]}}
Test.IsOk InStr(TestResult, "code") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsOk JsonResult("code") * 1 >= 200000
Test.IsOk JsonResult("code") * 1 <= 200004

End Sub

Function PublicKucoin(Method As String, ReqType As String, Optional ParamDict As Dictionary) As String

Dim url As String
PublicApiSite = "https://api.kucoin.com/api/v1"

MethodParams = DictToString(ParamDict, "URLENC")
If MethodParams <> "" Then MethodParams = "?" & MethodParams
urlPath = "/" & Method & MethodParams
url = PublicApiSite & urlPath

PublicKucoin = WebRequestURL(url, ReqType)

End Function
Function PrivateKucoin(Method As String, ReqType As String, Credentials As Dictionary, Optional ParamDict As Dictionary) As String

Dim NonceUnique As String
Dim url As String
Dim postdata As String

'Kucoin wants a 13-digit Nonce, use time correction if needed
NonceUnique = GetKucoinTime()

TradeApiSite = "https://api.kucoin.com"
ApiEndPoint = "/api/v1/" & Method

If ReqType = "GET" Or ReqType = "DELETE" Then
    'For GET, DELETE request, all query parameters need to be included in the request url. (e.g. /api/v1/accounts?currency=BTC)
    If Not ParamDict Is Nothing Then
        'OrderId -> add to URL
        For Each key In ParamDict.Keys
            If LCase(key) = "orderid" Then
                ApiEndPoint = ApiEndPoint & "/" & ParamDict(key)
                ParamDict.Remove key
                Exit For
            End If
        Next key
    End If
    
    MethodTxt = DictToString(ParamDict, "URLENC")
    If MethodTxt <> "" Then ApiEndPoint = ApiEndPoint & "?" & MethodTxt
    ReqBody = ""
Else
    'For POST, PUT request, all query parameters need to be included in the request body with JSON. (e.g. {"currency":"BTC"}). Do not include extra spaces in JSON strings.
    MethodTxt = ""
    ReqBody = DictToString(ParamDict, "JSON")
    postdata = ReqBody
End If

ApiForSign = NonceUnique & ReqType & ApiEndPoint & ReqBody
APIsign = ComputeHash_C("SHA256", ApiForSign, Credentials("secretKey"), "STR64")

url = TradeApiSite & ApiEndPoint

Dim headerDict As New Dictionary
headerDict.Add "KC-API-KEY", Credentials("apiKey")
headerDict.Add "KC-API-SIGN", APIsign
headerDict.Add "KC-API-TIMESTAMP", NonceUnique
headerDict.Add "KC-API-PASSPHRASE", Credentials("Passphrase")
headerDict.Add "Content-Type", "application/json"

PrivateKucoin = WebRequestURL(url, ReqType, headerDict, postdata)

End Function

Function GetKucoinTime() As Double

Dim JsonResponse As String
Dim Json As Object

'PublicKucoin time
JsonResponse = PublicKucoin("timestamp", "GET")
Set Json = JsonConverter.ParseJson(JsonResponse)
GetKucoinTime = Json("data")
If GetKucoinTime = 0 Then
    TimeCorrection = -3600
    GetKucoinTime = DateDiff("s", "1/1/1970", Now)
    GetKucoinTime = Trim(Str((Val(GetKucoinTime) + TimeCorrection)) & Right(Int(Timer * 100), 2) & "0")
End If

Set Json = Nothing

End Function
