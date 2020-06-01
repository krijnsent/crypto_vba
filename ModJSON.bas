Attribute VB_Name = "ModJSON"
'Functions in module:
'MaxDepth - integer with the maximum depth of the JSON
'JsonToArray - transforms JSON into an array with an internal tree structure
'ArrayTable - transforms JsonToArray (internal tree) into a flat table for output
'Source: https://github.com/krijnsent/crypto_vba

Sub TestJson()

Dim JsonResponse As String
Dim Json As Object 'Can be dictionary - json starting {} or collection - json starting []
Dim JsonRes As Dictionary
' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModJSON"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestDepth")
'Kraken Time
JsonResponse = "{""error"":[],""result"":{""unixtime"":1495455831,""rfc1123"":""Mon, 22 May 17 12:23:51 +0000""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
TestResult = MaxDepth(JsonRes)
Test.IsEqual TestResult, 1

'Poloniex returnTicker
JsonResponse = "{""BTC_BCN"":{""id"":7,""last"":""0.00000210"",""lowestAsk"":""0.00000210"",""highestBid"":""0.00000208"",""percentChange"":""0.73553719"",""baseVolume"":""26784.80209760"",""quoteVolume"":""13894501407.13100815"",""isFrozen"":""0"",""high24hr"":""0.00000280"",""low24hr"":""0.00000118""},""BTC_DASH"":{""id"":24,""last"":""0.04775443"",""lowestAsk"":""0.04781078"",""highestBid"":""0.04775443"",""percentChange"":""0.00446825"",""baseVolume"":""2884.45152468"",""quoteVolume"":""60634.59565660"",""isFrozen"":""0"",""high24hr"":""0.05035290"",""low24hr"":""0.04430738""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
TestResult = MaxDepth(Json)
Test.IsEqual TestResult, 2

'Kraken OHLC
JsonResponse = "{""error"":[],""result"":{""XXBTZEUR"":[[1492606800,""1121.990"",""1124.912"",""1119.680"",""1124.912"",""1122.345"",""352.76808800"",602],[1492610400,""1124.499"",""1124.980"",""1119.680"",""1122.000"",""1122.194"",""218.62127780"",713],[1492614000,""1121.311"",""1122.900"",""1120.501"",""1122.899"",""1122.266"",""445.46426003"",851],[1492617600,""1122.894"",""1124.499"",""1120.710"",""1123.291"",""1123.068"",""253.55336370"",860],[1492621200,""1124.406"",""1126.000"",""1123.017"",""1125.990"",""1124.775"",""234.27612705"",918],[1492624800,""1125.610"",""1126.231"",""1123.010"",""1126.229"",""1125.453"",""243.42246123"",772]],""last"":1495191600}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
TestResult = MaxDepth(JsonRes)
Test.IsEqual TestResult, 3

'WEXnz depth
JsonResponse = "{""btc_eur"":{""asks"":[[1919.99999,0.1111724],[1920,0.30236723],[1924.41,0.00601202],[1924.41522,0.009536]]}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
TestResult = MaxDepth(Json)
Test.IsEqual TestResult, 4


'TestJsonToArray
Set Test = Suite.Test("TestJsonToArray")
'Kraken Time
JsonResponse = "{""error"":[],""result"":{""unixtime"":1495455831,""rfc1123"":""Mon, 22 May 17 12:23:51 +0000""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
TestResult = JsonToArray(JsonRes)
Test.IsEqual UBound(TestResult, 1), 5
Test.IsEqual UBound(TestResult, 2), 3
Test.IsEqual TestResult(3, 2), "unixtime"
Test.IsEqual TestResult(3, 3), "rfc1123"

'Poloniex returnTicker
JsonResponse = "{""BTC_BCN"":{""id"":7,""last"":""0.00000210"",""lowestAsk"":""0.00000210"",""highestBid"":""0.00000208"",""percentChange"":""0.73553719"",""baseVolume"":""26784.80209760"",""quoteVolume"":""13894501407.13100815"",""isFrozen"":""0"",""high24hr"":""0.00000280"",""low24hr"":""0.00000118""},""BTC_DASH"":{""id"":24,""last"":""0.04775443"",""lowestAsk"":""0.04781078"",""highestBid"":""0.04775443"",""percentChange"":""0.00446825"",""baseVolume"":""2884.45152468"",""quoteVolume"":""60634.59565660"",""isFrozen"":""0"",""high24hr"":""0.05035290"",""low24hr"":""0.04430738""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
TestResult = JsonToArray(Json)
Test.IsEqual UBound(TestResult, 1), 5
Test.IsEqual UBound(TestResult, 2), 23
Test.IsEqual TestResult(3, 11), "high24hr"
Test.IsEqual TestResult(4, 14), 24

'Kraken OHLC
JsonResponse = "{""error"":[],""result"":{""XXBTZEUR"":[[1492606800,""1121.990"",""1124.912"",""1119.680"",""1124.912"",""1122.345"",""352.76808800"",602],[1492610400,""1124.499"",""1124.980"",""1119.680"",""1122.000"",""1122.194"",""218.62127780"",713],[1492614000,""1121.311"",""1122.900"",""1120.501"",""1122.899"",""1122.266"",""445.46426003"",851],[1492617600,""1122.894"",""1124.499"",""1120.710"",""1123.291"",""1123.068"",""253.55336370"",860],[1492621200,""1124.406"",""1126.000"",""1123.017"",""1125.990"",""1124.775"",""234.27612705"",918],[1492624800,""1125.610"",""1126.231"",""1123.010"",""1126.229"",""1125.453"",""243.42246123"",772]],""last"":1495191600}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
TestResult = JsonToArray(JsonRes)
Test.IsEqual UBound(TestResult, 1), 5
Test.IsEqual UBound(TestResult, 2), 57
Test.IsEqual TestResult(3, 11), 8
Test.IsEqual TestResult(4, 14), "1124.499"

'BTCe depth
JsonResponse = "{""btc_eur"":{""asks"":[[1919.99999,0.1111724],[1920,0.30236723],[1924.41,0.00601202],[1924.41522,0.009536]]}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
TestResult = JsonToArray(Json)
Test.IsEqual UBound(TestResult, 1), 5
Test.IsEqual UBound(TestResult, 2), 15
Test.IsEqual TestResult(3, 11), 1
Test.IsEqual TestResult(4, 14), 1924.41522


'TestArrayTable
Set Test = Suite.Test("TestArrayTable")

'Kraken Time
JsonResponse = "{""error"":[],""result"":{""unixtime"":1495455831,""rfc1123"":""Mon, 22 May 17 12:23:51 +0000""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
ResArr = JsonToArray(JsonRes)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 2
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(1, 1), "unixtime"
Test.IsEqual TestResult(2, 2), "Mon, 22 May 17 12:23:51 +0000"

'Poloniex returnTicker
JsonResponse = "{""BTC_BCN"":{""id"":7,""last"":""0.00000210"",""lowestAsk"":""0.00000210"",""highestBid"":""0.00000208"",""percentChange"":""0.73553719"",""baseVolume"":""26784.80209760"",""quoteVolume"":""13894501407.13100815"",""isFrozen"":""0"",""high24hr"":""0.00000280"",""low24hr"":""0.00000118""},""BTC_DASH"":{""id"":24,""last"":""0.04775443"",""lowestAsk"":""0.04781078"",""highestBid"":""0.04775443"",""percentChange"":""0.00446825"",""baseVolume"":""2884.45152468"",""quoteVolume"":""60634.59565660"",""isFrozen"":""0"",""high24hr"":""0.05035290"",""low24hr"":""0.04430738""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 11
Test.IsEqual UBound(TestResult, 2), 3
Test.IsEqual TestResult(1, 2), "BTC_BCN"
Test.IsEqual TestResult(3, 3), "0.04775443"

'Kraken OHLC
JsonResponse = "{""error"":[],""result"":{""XXBTZEUR"":[[1492606800,""1121.990"",""1124.912"",""1119.680"",""1124.912"",""1122.345"",""352.76808800"",602],[1492610400,""1124.499"",""1124.980"",""1119.680"",""1122.000"",""1122.194"",""218.62127780"",713],[1492614000,""1121.311"",""1122.900"",""1120.501"",""1122.899"",""1122.266"",""445.46426003"",851],[1492617600,""1122.894"",""1124.499"",""1120.710"",""1123.291"",""1123.068"",""253.55336370"",860],[1492621200,""1124.406"",""1126.000"",""1123.017"",""1125.990"",""1124.775"",""234.27612705"",918],[1492624800,""1125.610"",""1126.231"",""1123.010"",""1126.229"",""1125.453"",""243.42246123"",772]],""last"":1495191600}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 11
Test.IsEqual UBound(TestResult, 2), 7
Test.IsEqual TestResult(1, 2), "result"
Test.IsEqual TestResult(4, 4), 1492614000

'BTCe depth
JsonResponse = "{""btc_eur"":{""asks"":[[1919.99999,0.1111724],[1920,0.30236723],[1924.41,0.00601202],[1924.41522,0.009536]]}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 5
Test.IsEqual UBound(TestResult, 2), 5
Test.IsEqual TestResult(1, 2), "btc_eur"
Test.IsEqual TestResult(4, 4), 1924.41

'Poloniex deposit/withdrawal, no header output
JsonResponse = "{""deposits"":[{""currency"":""BTC"",""address"":""DEP1"",""amount"":""0.01006132"",""confirmations"":10,""txid"":""17f819a91369a9ff6c4a34216d434597cfc1b4a3d0489b46bd6f924137a47701"",""timestamp"":1399305798,""status"":""COMPLETE""},{""currency"":""BTC"",""address"":""DEP2"",""amount"":""0.00404104"",""confirmations"":10,""txid"":""7acb90965b252e55a894b535ef0b0b65f45821f2899e4a379d3e43799604695c"",""timestamp"":1399245916,""status"":""COMPLETE""}],""withdrawals"":[{""withdrawalNumber"":134933,""currency"":""BTC"",""address"":""1N2i5n8DwTGzUq2Vmn9TUL8J1vdr1XBDFg"",""amount"":""5.00010000"", ""timestamp"":1399267904,""status"":""COMPLETE: 36e483efa6aff9fd53a235177579d98451c4eb237c210e66cd2b9a2d4a988f8e"",""ipAddress"":""IP192""}]}"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, False)
Test.IsEqual UBound(TestResult, 1), 11
Test.IsEqual UBound(TestResult, 2), 3
Test.IsEqual TestResult(1, 2), "deposits"
Test.IsEqual TestResult(4, 2), "DEP2"

'Test no header reply
JsonResponse = "{""error"":[],""result"":{""XXBTZEUR"":[[1492606800,""1121.990"",""1124.912"",""1119.680"",""1124.912"",""1122.345"",""352.76808800"",602],[1492610400,""1124.499"",""1124.980"",""1119.680"",""1122.000"",""1122.194"",""218.62127780"",713],[1492614000,""1121.311"",""1122.900"",""1120.501"",""1122.899"",""1122.266"",""445.46426003"",851],[1492617600,""1122.894"",""1124.499"",""1120.710"",""1123.291"",""1123.068"",""253.55336370"",860],[1492621200,""1124.406"",""1126.000"",""1123.017"",""1125.990"",""1124.775"",""234.27612705"",918],[1492624800,""1125.610"",""1126.231"",""1123.010"",""1126.229"",""1125.453"",""243.42246123"",772]],""last"":1495191600}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, False)
Test.IsEqual UBound(TestResult, 1), 11
Test.IsEqual UBound(TestResult, 2), 6
Test.IsEqual TestResult(1, 2), "result"
Test.IsEqual TestResult(4, 2), 1492610400

'Empty data set returned 1
JsonResponse = "{""success"":true,""message"":"""",""result"":[]}"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 3
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(1, 2), True
Test.IsEqual TestResult(3, 2), 0

'Empty data set returned 2
JsonResponse = "{""success"":false,""message"":""APISIGN_NOT_PROVIDED"",""result"":null}"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 3
Test.IsEqual UBound(TestResult, 2), 2
Test.IsEqual TestResult(1, 2), False
Test.IsEqual TestResult(2, 2), "APISIGN_NOT_PROVIDED"

'Error set - only if Json is defined as Dictionary, as Object is okay
JsonResponse = "[{""balance"":0,""pendingFunds"":0,""currency"":""BCH""},{""balance"":41,""pendingFunds"":0,""currency"":""AUD""},{""balance"":145,""pendingFunds"":0,""currency"":""BTC""},{""balance"":0,""pendingFunds"":0,""currency"":""LTC""}]"
Set Json = JsonConverter.ParseJson(JsonResponse)
ResArr = JsonToArray(Json)
TestResult = ArrayTable(ResArr, True)
Test.IsEqual UBound(TestResult, 1), 4
Test.IsEqual UBound(TestResult, 2), 5
Test.IsEqual TestResult(2, 3), 41
Test.IsEqual TestResult(4, 4), "BTC"

'NEW TEST, UNFINISHED
PrTxt = "{""ret_msg"": ""ok"",""ext_code"": """",""result"": {""BTCUSD"": {""leverage"": 1},""EOSUSD"": {""leverage"": 1}},""time_now"": ""1567608910.732004""}"
Set Json = JsonConverter.ParseJson(PrTxt)
Set Res1 = Json("result")
'Res2 = json("result") '-> error, doesn't work...
For Each elm In Json("result")
    'Debug.Print elm
    'Debug.Print json("result")(elm)
    Set Res3 = Json("result")(elm)
    'Debug.Print json("result")(elm)("leverage")
Next elm

End Sub

Function MaxDepth(ObjIn As Object, Optional MaxLvl As Integer = 1, Optional NodeLvl As Integer = 1) As Integer
    
    Dim CollIn As New Collection
    Dim DictIn As New Scripting.Dictionary
    Dim iO As Object
    Dim Lvl As Integer
    
    If TypeName(ObjIn) = "Collection" Then
        'arrays ([]) to collections, arrays only have values
        Set CollIn = ObjIn
        For i = 1 To CollIn.Count
            'item could be value, object or array, determine:
            Set iO = Nothing
            On Error Resume Next
            Set iO = CollIn(i)
            On Error GoTo 0

            'item/value
            If Not (iO Is Nothing) Then
                If NodeLvl + 1 > MaxLvl Then MaxLvl = NodeLvl + 1
                NextLvl = MaxDepth(iO, MaxLvl, NodeLvl + 1)
            End If
        Next i
    ElseIf TypeName(ObjIn) = "Dictionary" Then
        'objects ({}) to dictionaries, Objects have key:values
        Set DictIn = ObjIn
        For Each k In DictIn.Keys
            'item could be value, object or array, determine:
            IV = ""
            Set iO = Nothing
            On Error Resume Next
            IV = DictIn(k)
            Set iO = DictIn(k)
            On Error GoTo 0
            
            'item/value
            If Not (iO Is Nothing) Then
                If NodeLvl + 1 > MaxLvl Then MaxLvl = NodeLvl + 1
                NextLvl = MaxDepth(iO, MaxLvl, NodeLvl + 1)
            End If
        Next k
    End If
    
    MaxDepth = MaxLvl
    
End Function



Function JsonToArray(ObjIn As Object, Optional ParentKey As String = "MAIN", Optional NodeLvl As Integer = 1, Optional ResArr As Variant) As Variant
    'Dim TempResArr() As Variant
    
    If IsMissing(ResArr) Then
        ReDim ResArr(1 To 5, 1 To 1)
        ResArr(1, 1) = "NODE_LVL"
        ResArr(2, 1) = "PARENT"
        ResArr(3, 1) = "KEY"
        ResArr(4, 1) = "VALUE"
        ResArr(5, 1) = "TYPE"
    End If

    Dim CollIn As New Collection
    Dim DictIn As New Scripting.Dictionary
    Dim iO As Object
    Dim CurK As String
    Dim Lvl As Integer
    
    If TypeName(ObjIn) = "Collection" Then
        'arrays ([]) to collections, arrays only have values
        Set CollIn = ObjIn
        For i = 1 To CollIn.Count
            'item could be value, object or array, determine:
            IV = ""
            Set iO = Nothing
            On Error Resume Next
            IV = CollIn(i)
            Set iO = CollIn(i)
            On Error GoTo 0

            'item/value
            If Not (iO Is Nothing) Then
                'Collection or Array, store and go one level deeper
                ReDim Preserve ResArr(1 To 5, 1 To UBound(ResArr, 2) + 1)
                ResArr(1, UBound(ResArr, 2)) = NodeLvl
                ResArr(2, UBound(ResArr, 2)) = ParentKey
                ResArr(3, UBound(ResArr, 2)) = i
                ResArr(4, UBound(ResArr, 2)) = iO.Count
                ResArr(5, UBound(ResArr, 2)) = "OBJ"
                'Debug.Print "LVL: " & NodeLvl & ", PARENT: " & ParentKey & " , KEY: " & I & " VALUE: count: " & iO.Count & " , TYPE:OBJ"
                ParentKey = i
                NextLvl = JsonToArray(iO, Str(i), NodeLvl + 1, ResArr)
            Else
                'item, write simple value
                'Debug.Print "LVL: " & NodeLvl & ", PARENT: " & ParentKey & " , KEY: " & I & " VALUE:" & iV & " , TYPE:VAL"
                ReDim Preserve ResArr(1 To 5, 1 To UBound(ResArr, 2) + 1)
                ResArr(1, UBound(ResArr, 2)) = NodeLvl
                ResArr(2, UBound(ResArr, 2)) = ParentKey
                ResArr(3, UBound(ResArr, 2)) = i
                ResArr(4, UBound(ResArr, 2)) = IV
                ResArr(5, UBound(ResArr, 2)) = "VAL"
            End If
        Next i
    ElseIf TypeName(ObjIn) = "Dictionary" Then
        'objects ({}) to dictionaries, Objects have key:values
        Set DictIn = ObjIn
        For Each k In DictIn.Keys
            'item could be value, object or array, determine:
            IV = ""
            Set iO = Nothing
            On Error Resume Next
            IV = DictIn(k)
            Set iO = DictIn(k)
            On Error GoTo 0
            
            'item/value
            If Not (iO Is Nothing) Then
                'Collection or Array, store and go one level deeper
                'Debug.Print "LVL: " & NodeLvl & ", PARENT: " & ParentKey & " , KEY: " & k & " VALUE: count: " & iO.Count & " , TYPE:OBJ"
                ReDim Preserve ResArr(1 To 5, 1 To UBound(ResArr, 2) + 1)
                ResArr(1, UBound(ResArr, 2)) = NodeLvl
                ResArr(2, UBound(ResArr, 2)) = ParentKey
                ResArr(3, UBound(ResArr, 2)) = k
                ResArr(4, UBound(ResArr, 2)) = iO.Count
                ResArr(5, UBound(ResArr, 2)) = "OBJ"
                CurK = k
                NextLvl = JsonToArray(iO, CurK, NodeLvl + 1, ResArr)
            Else
                'item, write simple value
                'Debug.Print "LVL: " & NodeLvl & ", PARENT: " & ParentKey & " , KEY: " & k & " VALUE:" & iV & " , TYPE:VAL"
                ReDim Preserve ResArr(1 To 5, 1 To UBound(ResArr, 2) + 1)
                ResArr(1, UBound(ResArr, 2)) = NodeLvl
                ResArr(2, UBound(ResArr, 2)) = ParentKey
                ResArr(3, UBound(ResArr, 2)) = k
                ResArr(4, UBound(ResArr, 2)) = IV
                ResArr(5, UBound(ResArr, 2)) = "VAL"
            End If
        Next k
    End If
    
    JsonToArray = ResArr

End Function
Function ArrayTable(ArrIn As Variant, Optional ReturnHeader As Boolean = True) As Variant

'Expected input: NODE_LVL -- PARENT -- KEY -- VALUE -- TYPE
Dim NrIt As Integer
Dim MaxD As Integer
Dim TblHeaders As New Scripting.Dictionary

'Get max depth and max items at that level
MaxD = 0
'Find maximum depth
For rw = LBound(ArrIn, 2) To UBound(ArrIn, 2)
    If Val(ArrIn(1, rw)) > MaxD Then
        MaxD = ArrIn(1, rw)
    End If
Next

'Get unique headers
On Error Resume Next
For rw = LBound(ArrIn, 2) To UBound(ArrIn, 2)
    Lvl = Val(ArrIn(1, rw))
    If Lvl < MaxD And Lvl > 0 Then
        TblHeaders.Add "GROUP_" & Lvl, "GROUP_" & Lvl
    'ElseIf Lvl = MaxD And ArrIn(5, rw) = "VAL" Then
    ElseIf Lvl = MaxD Then
        If Val(ArrIn(3, rw)) > 0 Then
            TblHeaders.Add "VAL_" & ArrIn(3, rw), "VAL_" & ArrIn(3, rw)
        Else
            TblHeaders.Add ArrIn(3, rw), ArrIn(3, rw)
        End If
    End If
Next
On Error GoTo 0

If ReturnHeader = True Then
    HeadRw = 1
Else
    HeadRw = 0
End If
ReDim ResArr(1 To TblHeaders.Count, 1 To 1 + HeadRw)

TempRw = 0
ResRw = 1 + HeadRw

For rw = LBound(ArrIn, 2) To UBound(ArrIn, 2)
    Lvl = Val(ArrIn(1, rw))
    If rw < UBound(ArrIn, 2) Then
        NextLvl = Val(ArrIn(1, rw + 1))
    Else
        NextLvl = 0
    End If
    If Lvl = MaxD Then
        'Get result column
        Idx = 0
        If Val(ArrIn(3, rw)) > 0 Then
            Idx = Application.Match("VAL_" & ArrIn(3, rw), TblHeaders.Keys, 0)
            If ReturnHeader = True Then
                ResArr(Idx, 1) = "VAL_" & ArrIn(3, rw)
            End If
        Else
            Idx = Application.Match(ArrIn(3, rw), TblHeaders.Keys, 0)
            If ReturnHeader = True Then
                ResArr(Idx, 1) = ArrIn(3, rw)
            End If
        End If
        
        ResArr(Idx, ResRw) = ArrIn(4, rw)
        For k = 1 To Lvl
            If IsEmpty(ResArr(k, ResRw)) Then ResArr(k, ResRw) = ResArr(k, ResRw - 1)
        Next k
        TempRw = TempRw + 1
        If rw < UBound(ArrIn, 2) And NextLvl < Lvl Then
            TempRw = 0
            ResRw = ResRw + 1
            ReDim Preserve ResArr(1 To TblHeaders.Count, 1 To ResRw)
        End If
    ElseIf Lvl > 0 Then
        If ReturnHeader = True Then
            ResArr(Lvl, 1) = "GROUP_" & Lvl
        End If
        ResArr(Lvl, ResRw) = ArrIn(3, rw)
    End If
Next

'Strip last line if that wasn't a max depth record
If Lvl < MaxD Then
    ReDim Preserve ResArr(1 To TblHeaders.Count, 1 To ResRw - 1)
End If
ArrayTable = ResArr

End Function
