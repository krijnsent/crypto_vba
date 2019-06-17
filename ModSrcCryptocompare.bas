Attribute VB_Name = "ModSrcCryptocompare"
'Two variables for caching, so the formulas don't update every recalculation
Public Const CCCacheSeconds = 60   'Nr of seconds cache, default >= 60
Public CCDict As New Scripting.Dictionary

Sub TestSrcCryptocompare()

'This module contains functions to use in a sheet or in VBA
'Source: https://github.com/krijnsent/crypto_vba
'Note: the functions currently slow down the sheets massively, use max 10 functions per workbook, otherwise your workbook might CRASH
'ToDo: better error catching
'For cryptocompare, please get a free API key at https://www.cryptocompare.com

'Functions in this module:
'C_LAST_PRICE - price?fsym=BTC&tsyms=USD,EUR&e=Coinbase
'C_HIST_PRICE - pricehistorical?fsym=BTC&tsyms=USD,EUR&e=Coinbase&ts=1452680400
'C_DAY_AVG_PRICE - dayAvg?fsym=BTC&tsym=USD&toTs=1487116800&e=Bitfinex
'C_ARR_OHLCV - histoday?fsym=GBP&tsym=USD&limit=30&aggregate=1&e=CCCAGG

Dim Apikey As String
Apikey = "your_api_key_here" 'empty if you don't use an API key

'Remove this line, unless you define a constant somewhere ( Public Const apikey_cryptocompare = "the key to use everywhere" etc )
Apikey = apikey_cryptocompare

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModSrcCryptocompare"
' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestPublicCryptoCompareData")

'Error, unknown command
TestResult = PublicCryptoCompareData("unknown_command")
Test.IsOk InStr(TestResult, "Error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "Path does not exist"
Test.IsEqual JsonResult("Path"), ""

'Error, no parameters
TestResult = PublicCryptoCompareData("data/histoday")
Test.IsOk InStr(TestResult, "Error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "fsym param seems to be missing."
Test.IsEqual JsonResult("Path"), ""

'Error, create a dictionary with ONLY the parameter fsym
Dim Params As New Dictionary
Params.Add "fsym", "BTC"
TestResult = PublicCryptoCompareData("data/histoday", Params)
Test.IsOk InStr(TestResult, "Error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "tsym param seems to be missing."
Test.IsEqual JsonResult("Path"), ""

'Error, add to the same dictionary an unknown tsym
Params.Add "tsym", "BLABLA"
TestResult = PublicCryptoCompareData("data/histoday", Params)
Test.IsOk InStr(TestResult, "Error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "There is no data for the toSymbol BLABLA ."
Test.IsEqual JsonResult("Path"), ""

'OK, two correct parameters for histoday, make a new dictionary
Dim Params2 As New Dictionary
Params2.Add "fsym", "BTC"
Params2.Add "tsym", "XMR"
TestResult = PublicCryptoCompareData("data/histoday", Params2)
Test.IsOk InStr(TestResult, "Success") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Success"
Test.IsEqual JsonResult("Message"), ""
Test.IsEqual JsonResult("Path"), ""
Test.IsOk JsonResult("Data")(1)("high") > 0

'Tests for APIkey (please get a free key at Cryptocompare.com)
'This test fails, as API key is needed:
TestResult = PublicCryptoCompareData("data/social/coin/latest")
'{"Response":"Error","Message":"You need a valid auth key or api key to access this endpoint","HasWarning":false,"Type":1,"RateLimit":{},"Data":{}}
Test.IsOk InStr(TestResult, "Error") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "You need a valid auth key or api key to access this endpoint"
Test.IsEqual JsonResult("Type"), 1

'Add an API key and force caching off
Dim Params3 As New Dictionary
Params3.Add "apikey", Apikey
TestResult = PublicCryptoCompareData("data/social/coin/latest", Params3)
'{"Response":"Success","Message":"","HasWarning":false,"Type":100,"RateLimit":{},"Data":{"General":{"Points":8212774,"Name":"BTC","CoinName":"Bitcoin","Type":"Webpagecoinp"},"CryptoCompare":{"Points":6898505, etc...
Test.IsOk InStr(TestResult, "Success") > 0
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Success"
Test.IsEqual JsonResult("Message"), ""
Test.IsEqual JsonResult("Type"), 100
If JsonResult("Type") = 100 Then
    Test.IsEqual JsonResult("Data")("General")("Name"), "BTC"
End If

'Rate limit WITHOUT an API key: 1000/minute
TestResult = PublicCryptoCompareData("stats/rate/limit")
'{"Response":"Success","Message":"","HasWarning":false,"Type":100,"RateLimit":{},"Data":{"calls_made":{"second":1,"minute":10,"hour":138,"day":475,"month":4113},"calls_left":{"second":49,"minute":990,"hour":19862,"day":199525,"month":1995887}}}
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Success"
If JsonResult("Response") = "Success" Then
    Test.IsEqual JsonResult("Data")("calls_made")("minute") + JsonResult("Data")("calls_left")("minute"), 1000
End If

'Rate limit WITH an API key: 2500/minute
TestResult = PublicCryptoCompareData("stats/rate/limit", Params3)
'{"Response":"Success","Message":"","HasWarning":false,"Type":100,"RateLimit":{},"Data":{"calls_made":{"second":2,"minute":2,"hour":21,"day":33,"month":33},"calls_left":{"second":48,"minute":2498,"hour":24979,"day":49967,"month":99967}}}
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("Response"), "Success"
If JsonResult("Response") = "Success" Then
    Test.IsEqual JsonResult("Data")("calls_made")("minute") + JsonResult("Data")("calls_left")("minute"), 2500
End If


Set Test = Suite.Test("TestC_LAST_PRICE")
JsonResult = C_LAST_PRICE("MYCOIN1", "BLABLA")
Test.IsEqual JsonResult, "ERROR There is no data for the symbol MYCOIN1 ."

JsonResult = C_LAST_PRICE("BTC", "BLABLA")
Test.IsEqual JsonResult, "ERROR There is no data for any of the toSymbols BLABLA ."

JsonResult = C_LAST_PRICE("BTC", "EUR", "An_Unknown_Exchange")
Test.IsEqual JsonResult, "ERROR an_unknown_exchange market does not exist for this coin pair (BTC-EUR)"

JsonResult = C_LAST_PRICE("BTC", "EUR")
Test.IsOk JsonResult > 0

JsonResult = C_LAST_PRICE("BTC", "EUR", "Kraken")
Test.IsOk JsonResult > 0

'Optional, add an apikey, only affects the rate limit for this function
JsonResult = C_LAST_PRICE("BTC", "USD", "Bittrex", Apikey)
Test.IsOk JsonResult > 0


Set Test = Suite.Test("TestC_HIST_PRICE")
JsonResult = C_HIST_PRICE("ETH", "USD", "2018-01-01 20:00")
Test.IsOk JsonResult > 0
JsonResult = C_HIST_PRICE("ETH", "USD", #1/1/2018#, "Bittrex")
Test.IsOk JsonResult > 0
'Optional, add an apikey, only affects the rate limit for this function
JsonResult = C_HIST_PRICE("ETH", "USD", #1/1/2019#, , Apikey)
Test.IsOk JsonResult > 0


Set Test = Suite.Test("TestC_DAY_AVG_PRICE")
JsonResult = C_DAY_AVG_PRICE("ETH", "BTC", #1/1/2017#)
Test.IsOk JsonResult > 0
JsonResult = C_DAY_AVG_PRICE("ETH", "BTC", #1/1/2017#, "Poloniex")
Test.IsOk JsonResult > 0
'Optional, add an apikey, only affects the rate limit for this function
JsonResult = C_DAY_AVG_PRICE("XMR", "BTC", #10/1/2018#, , Apikey)
Test.IsOk JsonResult > 0


Set Test = Suite.Test("TestC_ARR_OHLCV")
'Function C_ARR_OHLCV(
'DayHour as String, CurrBuy As String, CurrSell As String, ReturnColumns As String -> ETCHLOFV, Optional NrHours As Long,
'Optional MaxTimeDate As Date, Optional Exchange As String, Optional exchange As String, Optional ReverseData As Boolean,
'Optional Apikey As String) As Variant()

'Test for errors first
TestArr = C_ARR_OHLCV("A", "2FA", "EUR", "ECV")
Test.IsEqual TestArr(1, 1), "ERROR, DayHourMin must end with D, H or M"

TestArr = C_ARR_OHLCV("90M", "2FA", "EUR", "ECV")
Test.IsEqual TestArr(1, 1), "ERROR, DayHourMin aggregation has to be from 1 to 60. Valid values are e.g. 7D, 2H or 30M"

TestArr = C_ARR_OHLCV("H", "ETH", "EUR", "")
Test.IsEqual TestArr(1, 1), "ERROR ReturnColumns, use the letters ETCHLOFV"

TestArr = C_ARR_OHLCV("H", "BTC", "EUR", "ABD")
Test.IsEqual TestArr(1, 1), "unknown ReturnColumn"
Test.IsEqual TestArr(1, 2), "unknown ReturnColumn"
Test.IsEqual TestArr(1, 3), "unknown ReturnColumn"

TestArr = C_ARR_OHLCV("H", "2FA", "EUR", "ECV")
Test.IsEqual TestArr(1, 1), "ERROR There is no data for the symbol 2FA ."

TestArr = C_ARR_OHLCV("H", "BTC", "EUR", "TECV", 48, #1/1/2018#, "Kraken")
Test.IsEqual UBound(TestArr, 1), 50
Test.IsEqual UBound(TestArr, 2), 4
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 2), "time"
Test.IsEqual TestArr(1, 3), "close"
Test.IsEqual TestArr(1, 4), "volumeto"
Test.IsEqual TestArr(2, 1), "1514592000"
Test.IsEqual TestArr(2, 2), #12/30/2017#
Test.IsOk TestArr(2, 3) > 1
Test.IsOk TestArr(2, 4) > 1

TestArr = C_ARR_OHLCV("H", "BTC", "EUR", "EC", 24, , "Kraken")
Test.IsEqual UBound(TestArr, 1), 26
Test.IsEqual UBound(TestArr, 2), 2
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 2), "close"
Test.IsOk TestArr(2, 1) > #12/30/2017#
Test.IsOk TestArr(2, 2) > 1

TestArr = C_ARR_OHLCV("H", "XLM", "EUR", "TEOHLCFV", 48, DateSerial(2018, 1, 1), "Kraken")
Test.IsEqual UBound(TestArr, 1), 50
Test.IsEqual UBound(TestArr, 2), 8
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 8), "volumeto"
Test.IsEqual TestArr(2, 2), #12/30/2017#
Test.IsEqual TestArr(50, 2), #1/1/2018#
Test.IsEqual TestArr(50, 3), 0.01447

TestArr = C_ARR_OHLCV("4H", "XMR", "BTC", "EC", 48)
Test.IsEqual UBound(TestArr, 1), 50
Test.IsEqual UBound(TestArr, 2), 2
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 2), "close"
Test.IsOk TestArr(50, 1) > #1/1/2018#
Test.IsOk TestArr(50, 2) > 0

'Flip the result (newest row on top)
TestArr = C_ARR_OHLCV("H", "XLM", "EUR", "TEOHLCFV", 24, DateSerial(2019, 1, 1), "Kraken", True, Apikey)
Test.IsEqual UBound(TestArr, 1), 26
Test.IsEqual UBound(TestArr, 2), 8
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 8), "volumeto"
Test.IsEqual TestArr(2, 2), #1/1/2019#
Test.IsEqual TestArr(26, 2), #12/31/2018#
Test.IsEqual TestArr(26, 3), 0.1022

End Sub

Function PublicCryptoCompareData(Method As String, Optional ParamDict As Dictionary) As String

'For documentation, see: https://min-api.cryptocompare.com/
Dim Url As String
Dim Apikey As String
Dim TempData As String
Dim Sec As Double
Dim objHeaders As New Dictionary

PublicApiSite = "https://min-api.cryptocompare.com/"
'Check for API key and move that to the header of the GET request.
MethodParams = ""
If Not ParamDict Is Nothing Then
    If ParamDict.Exists("apikey") Then
        'move to the end
        tempkey = ParamDict("apikey")
        ParamDict.Remove "apikey"
        ParamDict.Add ("api_key"), tempkey
    End If
    'Change the rest of the parameters to JSON
    MethodParams = DictToString(ParamDict, "URLENC")
    If MethodParams <> "" Then MethodParams = "?" & MethodParams
End If

urlPath = Method & MethodParams
Url = PublicApiSite & urlPath

'For caching, check if data already exists
IsInDict = CCDict.Exists(urlPath)
GetNewData = False
If IsInDict = True Then
    'In dictionary, check time
    If CCDict(urlPath) + TimeSerial(0, 0, CCCacheSeconds) < Now() Then
        'Has not been updated recently and/or forced no caching, update now
        CCDict.Remove urlPath
        CCDict.Add urlPath, Now()
        If CCDict.Exists("DATA-" & urlPath) Then CCDict.Remove "DATA-" & urlPath
        GetNewData = True
    End If
Else
    CCDict.Add urlPath, Now()
    GetNewData = True
End If

If GetNewData = True Then
    TempData = WebRequestURL(Url, "GET", objHeaders)
    CCDict.Add "DATA-" & urlPath, TempData
Else
    TempData = CCDict("DATA-" & urlPath)
End If

PublicCryptoCompareData = TempData

End Function
Function C_LAST_PRICE(CurrBuy As String, CurrSell As String, Optional exchange As String, Optional Apikey As String)

Dim PrTxt As String
Dim json As Object
Dim ParamDict As New Dictionary
Application.Volatile

ParamDict.Add ("fsym"), CurrBuy
ParamDict.Add ("tsyms"), CurrSell
If Len(exchange) > 2 Then
    ParamDict.Add ("e"), exchange
End If
If Len(Apikey) > 0 Then
    ParamDict.Add ("apikey"), Apikey
End If

PrTxt = PublicCryptoCompareData("data/price", ParamDict)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_LAST_PRICE = "ERROR " & json("Message")
Else
    C_LAST_PRICE = json(CurrSell)
End If

Set json = Nothing

End Function

Function C_HIST_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional exchange As String, Optional Apikey As String)

Dim PrTxt As String
Dim json As Object
Dim ParamDict As New Dictionary
Application.Volatile

dt = DateToUnixTime(DateRates)
ParamDict.Add ("fsym"), CurrBuy
ParamDict.Add ("tsyms"), CurrSell
ParamDict.Add ("ts"), dt
If Len(exchange) > 2 Then
    ParamDict.Add ("e"), exchange
End If
If Len(Apikey) > 0 Then
    ParamDict.Add ("apikey"), Apikey
End If

PrTxt = PublicCryptoCompareData("data/price", ParamDict)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_HIST_PRICE = "ERROR " & json("Message")
Else
    C_HIST_PRICE = json(CurrSell)
End If

Set json = Nothing

End Function

Function C_DAY_AVG_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional exchange As String, Optional Apikey As String)

Dim PrTxt As String
Dim json As Object
Dim ParamDict As New Dictionary
Application.Volatile

dt = DateToUnixTime(DateRates)
ParamDict.Add ("fsym"), CurrBuy
ParamDict.Add ("tsym"), CurrSell
ParamDict.Add ("toTs"), dt
If Len(exchange) > 2 Then
    ParamDict.Add ("e"), exchange
End If
If Len(Apikey) > 0 Then
    ParamDict.Add ("apikey"), Apikey
End If

PrTxt = PublicCryptoCompareData("data/dayAvg", ParamDict)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_DAY_AVG_PRICE = "ERROR " & json("Message")
Else
    C_DAY_AVG_PRICE = json(CurrSell)
End If

Set json = Nothing

End Function

Function C_ARR_OHLCV(DayHourMin As String, CurrBuy As String, CurrSell As String, ReturnColumns As String, Optional NrLines As Long, Optional MaxTimeDate As Date, Optional exchange As String, Optional ReverseData As Boolean, Optional Apikey As String) As Variant()

'ReturnColumns: variable "TEOHLCFV" -> select columns you want back in the order you want them back, no spaces
'T = timestamp (unixtime)
'E = normal excel date/time
'O = open price
'H = high price
'L = Low price
'C = close price
'F = volume From
'V = volume to

Dim ExchangeTxt As String
Dim PrTxt As String
Dim AggrVal As String
Dim cmd As String
Dim utime As Long
Dim json As Object
Dim TempArr As Variant
Dim ParamDict As New Dictionary
ColumnOptions = "ETCHLOFV"
Application.Volatile

If UCase(Right(DayHourMin, 1)) = "D" Then
    cmd = "data/histoday"
ElseIf UCase(Right(DayHourMin, 1)) = "H" Then
    cmd = "data/histohour"
ElseIf UCase(Right(DayHourMin, 1)) = "M" Then
    cmd = "data/histominute"
Else
    'Error
    ReDim TempArr(1 To 1, 1 To 1)
    TempArr(1, 1) = "ERROR, DayHourMin must end with D, H or M"
    C_ARR_OHLCV = TempArr
    Exit Function
End If

ParamDict.Add ("fsym"), CurrBuy
ParamDict.Add ("tsym"), CurrSell
If Len(DayHourMin) > 1 Then
    AggrVal = Left(DayHourMin, Len(DayHourMin) - 1)
    If Val(AggrVal) >= 1 And Val(AggrVal) <= 60 Then
        ParamDict.Add ("aggregate"), Val(AggrVal)
    Else
        'Error
        ReDim TempArr(1 To 1, 1 To 1)
        TempArr(1, 1) = "ERROR, DayHourMin aggregation has to be from 1 to 60. Valid values are e.g. 7D, 2H or 30M"
        C_ARR_OHLCV = TempArr
        Exit Function
    End If
End If

If MaxTimeDate > DateSerial(2000, 1, 1) Then
    dt = DateToUnixTime(MaxTimeDate)
    ParamDict.Add ("toTs"), dt
End If

If Len(exchange) > 2 Then
    ParamDict.Add ("e"), exchange
End If
If NrLines > 0 Then
    ParamDict.Add ("limit"), NrLines
End If
If Len(Apikey) > 0 Then
    ParamDict.Add ("apikey"), Apikey
End If

PrTxt = PublicCryptoCompareData(cmd, ParamDict)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    ReDim TempArr(1 To 1, 1 To 1)
    TempArr(1, 1) = "ERROR " & json("Message")
    C_ARR_OHLCV = TempArr
Else
    If InStr(PrTxt, """Data"":[]") > 0 Then
        'Empty result from Cryptocompare API, show user error
        ReDim TempArr(1 To 1, 1 To 1)
        TempArr(1, 1) = "ERROR, cryptocompare API gave back an empty result, try other settings"
        C_ARR_OHLCV = TempArr
        Exit Function
    End If
    ResArr = JsonToArray(json)
    ResTbl = ArrayTable(ResArr, True)
    
    ReturnColumns = UCase(Trim(ReturnColumns))
    'Process all columns
    If Len(ReturnColumns) > 0 Then
        ReDim TempArr(1 To UBound(ResTbl, 2), 1 To Len(ReturnColumns))
        For i = 1 To Len(ReturnColumns)
            Itm = Mid(ReturnColumns, i, 1)
            itmnr = InStr(ColumnOptions, Itm) + 1
            'Checked for valid column types, move the data to the TempArr
            If itmnr > 1 Then
                For j = 1 To UBound(ResTbl, 2)
                    j2 = j
                    If ReverseData = True And j > 1 Then j2 = UBound(ResTbl, 2) - j + 2
                    TempArr(j2, i) = ResTbl(itmnr, j)
                    If itmnr = 2 Then
                        'Time from Unixtime to normal date/time
                        If j > 1 Then
                            utime = ResTbl(itmnr + 1, j)
                            TempArr(j2, i) = UnixTimeToDate(utime)
                        Else
                            TempArr(j2, i) = ResTbl(itmnr + 1, j)
                        End If
                    End If
                Next j
            Else
                'Unknown column, no data to return
                For j = 1 To UBound(ResTbl, 2)
                    TempArr(j, i) = "unknown ReturnColumn"
                Next j
            End If
        Next i
        C_ARR_OHLCV = TempArr
    Else
        'No returncolumns identified, return error
        ReDim TempArr(1 To 1, 1 To 1)
        TempArr(1, 1) = "ERROR ReturnColumns, use the letters " & ColumnOptions
        C_ARR_OHLCV = TempArr
    End If
End If

Set json = Nothing

End Function


