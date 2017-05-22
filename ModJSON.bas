Attribute VB_Name = "ModJSON"
Sub TestDepth()

Dim JsonResponse As String
Dim Json As Dictionary
Dim JsonRes As Dictionary

'Kraken Time
JsonResponse = "{""error"":[],""result"":{""unixtime"":1495455831,""rfc1123"":""Mon, 22 May 17 12:23:51 +0000""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
Debug.Print MaxDepth(JsonRes)
'1

'Poloniex returnTicker
JsonResponse = "{""BTC_BCN"":{""id"":7,""last"":""0.00000210"",""lowestAsk"":""0.00000210"",""highestBid"":""0.00000208"",""percentChange"":""0.73553719"",""baseVolume"":""26784.80209760"",""quoteVolume"":""13894501407.13100815"",""isFrozen"":""0"",""high24hr"":""0.00000280"",""low24hr"":""0.00000118""},""BTC_DASH"":{""id"":24,""last"":""0.04775443"",""lowestAsk"":""0.04781078"",""highestBid"":""0.04775443"",""percentChange"":""0.00446825"",""baseVolume"":""2884.45152468"",""quoteVolume"":""60634.59565660"",""isFrozen"":""0"",""high24hr"":""0.05035290"",""low24hr"":""0.04430738""}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Debug.Print MaxDepth(Json)
'2

'Kraken OHLC
JsonResponse = "{""error"":[],""result"":{""XXBTZEUR"":[[1492606800,""1121.990"",""1124.912"",""1119.680"",""1124.912"",""1122.345"",""352.76808800"",602],[1492610400,""1124.499"",""1124.980"",""1119.680"",""1122.000"",""1122.194"",""218.62127780"",713],[1492614000,""1121.311"",""1122.900"",""1120.501"",""1122.899"",""1122.266"",""445.46426003"",851],[1492617600,""1122.894"",""1124.499"",""1120.710"",""1123.291"",""1123.068"",""253.55336370"",860],[1492621200,""1124.406"",""1126.000"",""1123.017"",""1125.990"",""1124.775"",""234.27612705"",918],[1492624800,""1125.610"",""1126.231"",""1123.010"",""1126.229"",""1125.453"",""243.42246123"",772]],""last"":1495191600}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Set JsonRes = Json("result")
Debug.Print MaxDepth(JsonRes)
'3

'BTCe depth
JsonResponse = "{""btc_eur"":{""asks"":[[1919.99999,0.1111724],[1920,0.30236723],[1924.41,0.00601202],[1924.41522,0.009536]]}}"
Set Json = JsonConverter.ParseJson(JsonResponse)
Debug.Print MaxDepth(Json)
'4


End Sub

Function MaxDepth(ObjIn As Object, Optional MaxLvl As Integer = 1, Optional NodeLvl As Integer = 1) As Integer
    
    Dim CollIn As New Collection
    Dim DictIn As New Scripting.Dictionary
    Dim iO As Object
    Dim Lvl As Integer
    
    If TypeName(ObjIn) = "Collection" Then
        'arrays ([]) to collections, arrays only have values
        Set CollIn = ObjIn
        For I = 1 To CollIn.Count
            'item could be value, object or array, determine:
            Set iO = Nothing
            On Error Resume Next
            Set iO = CollIn(I)
            On Error GoTo 0

            'item/value
            If Not (iO Is Nothing) Then
                If NodeLvl + 1 > MaxLvl Then MaxLvl = NodeLvl + 1
                NextLvl = MaxDepth(iO, MaxLvl, NodeLvl + 1)
            End If
        Next I
    ElseIf TypeName(ObjIn) = "Dictionary" Then
        'objects ({}) to dictionaries, Objects have key:values
        Set DictIn = ObjIn
        For Each k In DictIn.keys
            'item could be value, object or array, determine:
            iV = ""
            Set iO = Nothing
            On Error Resume Next
            iV = DictIn(k)
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
