Attribute VB_Name = "ModSrcCryptocompare"
'For historical prices, include https://www.cryptocompare.com/api/
'CacheTime: the nr of seconds before the cache value is refreshed
'ToDo: include caching - https://fastexcel.wordpress.com/2012/12/05/writing-efficient-udfs-part-12-getting-used-range-fast-using-application-events-and-a-cache/
'And https://www.experts-exchange.com/articles/1135/Use-Excel's-hidden-data-store-to-share-data-across-VBA-projects.html
'And https://github.com/jbaurle/PMStockQuote/blob/master/PMStockQuote/UserDefinedFunctions.cs
Public Const CacheTime = 1000

'Functions to use in a sheet
'Source: https://github.com/krijnsent/crypto_vba
'Note: the functions currently slow down the sheets massively, use max 10 functions per workbook, otherwise your workbook might CRASH
'ToDo: better error catching

'Functions:
'C_LAST_PRICE -  price?fsym=BTC&tsyms=USD,EUR&e=Coinbase
'C_HIST_PRICE - pricehistorical?fsym=BTC&tsyms=USD,EUR&e=Coinbase&ts=1452680400
'C_DAY_AVG_PRICE - dayAvg?fsym=BTC&tsym=USD&toTs=1487116800&e=Bitfinex
'C_ARR_OHLCV - histoday?fsym=GBP&tsym=USD&limit=30&aggregate=1&e=CCCAGG

Sub TestSrcCryptocompare()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModSrcCryptocompare"
' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
' Create a new test
Dim Test As TestCase


Set Test = Suite.Test("TestPublicCryptoCompareData")

'Test for errors first
Set JsonResult = JsonConverter.ParseJson(PublicCryptoCompareData("unknown_command", ""))
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "Path does not exist"
Test.IsEqual JsonResult("Path"), ""

Set JsonResult = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", ""))
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "fsym param seems to be missing."
Test.IsEqual JsonResult("Path"), ""

'Error--fsym param seems to be missing.--
Set JsonResult = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "fsym=BTC"))
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "Path does not exist"
Test.IsEqual JsonResult("Path"), ""

Set JsonResult = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "?fsym=BTC&tsym=BLABLA"))
Test.IsEqual JsonResult("Response"), "Error"
Test.IsEqual JsonResult("Message"), "There is no data for the toSymbol BLABLA ."
Test.IsEqual JsonResult("Path"), ""

Set JsonResult = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "?fsym=BTC&tsym=XMR"))
Test.IsEqual JsonResult("Response"), "Success"
Test.IsEqual JsonResult("Message"), ""
Test.IsEqual JsonResult("Path"), ""


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


Set Test = Suite.Test("TestC_HIST_PRICE")

JsonResult = C_HIST_PRICE("ETH", "USD", "2018-01-01 20:00")
Test.IsOk JsonResult > 0
JsonResult = C_HIST_PRICE("ETH", "USD", #1/1/2018#, "Bittrex")
Test.IsOk JsonResult > 0


Set Test = Suite.Test("TestC_DAY_AVG_PRICE")

JsonResult = C_DAY_AVG_PRICE("ETH", "BTC", #1/1/2017#)
Test.IsOk JsonResult > 0
JsonResult = C_DAY_AVG_PRICE("ETH", "BTC", #1/1/2017#, "Poloniex")
Test.IsOk JsonResult > 0


Set Test = Suite.Test("TestC_ARR_OHLCV")
'Function C_ARR_OHLCV(
'DayHour as String, CurrBuy As String, CurrSell As String, ReturnColumns As String -> ETCHLOFV, Optional NrHours As Long,
'Optional MaxTimeDate As Date, Optional Exchange As String) As Variant()
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
Test.IsEqual TestArr(50, 2), #1/1/2018#
Test.IsEqual TestArr(50, 3), 0.01447

TestArr = C_ARR_OHLCV("4H", "XMR", "BTC", "EC", 48)
Test.IsEqual UBound(TestArr, 1), 50
Test.IsEqual UBound(TestArr, 2), 2
Test.IsEqual TestArr(1, 1), "time"
Test.IsEqual TestArr(1, 2), "close"
Test.IsOk TestArr(50, 1) > #1/1/2018#
Test.IsOk TestArr(50, 2) > 0

TestArr = C_ARR_OHLCV("BLA", "XLM", "EUR", "EC", 48)
Test.IsEqual TestArr(1, 1), "ERROR, DayHourMin must end with D, H or M"

TestArr = C_ARR_OHLCV("99H", "XMR", "BTC", "EC", 48)
Test.IsEqual TestArr(1, 1), "ERROR, DayHourMin aggregation has to be from 1 to 60. Valid values are e.g. 7D, 2H or 30M"

TestArr = C_ARR_OHLCV("HH", "XMR", "BTC", "EC", 48)
Test.IsEqual TestArr(1, 1), "ERROR, DayHourMin aggregation has to be from 1 to 60. Valid values are e.g. 7D, 2H or 30M"

TestArr = C_ARR_OHLCV("H", "BTC", "EUR", "ABD")
Test.IsEqual TestArr(1, 1), "unknown ReturnColumn"
Test.IsEqual TestArr(1, 2), "unknown ReturnColumn"
Test.IsEqual TestArr(1, 3), "unknown ReturnColumn"

TestArr = C_ARR_OHLCV("H", "2FA", "EUR", "ECV")
Test.IsEqual TestArr(1, 1), "ERROR There is no data for the symbol 2FA ."

TestArr = C_ARR_OHLCV("H", "ETH", "EUR", "")
Test.IsEqual TestArr(1, 1), "ERROR ReturnColumns, use the letters ETCHLOFV"


End Sub

Function PublicCryptoCompareData(Method As String, Optional MethodOptions As String) As String

'https://www.cryptocompare.com/api/ or https://min-api.cryptocompare.com/
Dim Url As String

PublicApiSite = "https://min-api.cryptocompare.com/data"
urlPath = "/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicCryptoCompareData = WebRequestURL(Url, "GET")

Set objHTTP = Nothing

End Function
Function C_LAST_PRICE(CurrBuy As String, CurrSell As String, Optional exchange As String)

Dim PrTxt As String
Dim json As Object
Application.Volatile

If Len(exchange) > 2 Then
    ExchangeTxt = "&e=" & exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("price", "?fsym=" & CurrBuy & "&tsyms=" & CurrSell & ExchangeTxt)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_LAST_PRICE = "ERROR " & json("Message")
Else
    C_LAST_PRICE = json(CurrSell)
End If

Set json = Nothing

End Function

Function C_HIST_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional exchange As String)

Dim PrTxt As String
Dim json As Object
Application.Volatile

dt = DateToUnixTime(DateRates)
If Len(exchange) > 2 Then
    ExchangeTxt = "&e=" & exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("pricehistorical", "?fsym=" & CurrBuy & "&tsyms=" & CurrSell & "&ts=" & dt & ExchangeTxt)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_HIST_PRICE = "ERROR " & json("Message")
Else
    C_HIST_PRICE = json(CurrBuy)(CurrSell)
End If

Set json = Nothing

End Function

Function C_DAY_AVG_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional exchange As String)

Dim PrTxt As String
Dim json As Object
Application.Volatile

dt = DateToUnixTime(DateRates)
If Len(exchange) > 2 Then
    ExchangeTxt = "&e=" & exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("dayAvg", "?fsym=" & CurrBuy & "&tsym=" & CurrSell & "&toTs=" & dt & ExchangeTxt)
Set json = JsonConverter.ParseJson(PrTxt)

If json("Response") = "Error" Then
    'Error
    C_DAY_AVG_PRICE = "ERROR " & json("Message")
Else
    C_DAY_AVG_PRICE = json(CurrSell)
End If

Set json = Nothing

End Function

Function C_ARR_OHLCV(DayHourMin As String, CurrBuy As String, CurrSell As String, ReturnColumns As String, Optional NrLines As Long, Optional MaxTimeDate As Date, Optional exchange As String) As Variant()

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
ColumnOptions = "ETCHLOFV"
Application.Volatile

If UCase(Right(DayHourMin, 1)) = "D" Then
    cmd = "histoday"
ElseIf UCase(Right(DayHourMin, 1)) = "H" Then
    cmd = "histohour"
ElseIf UCase(Right(DayHourMin, 1)) = "M" Then
    cmd = "histominute"
Else
    'Error
    ReDim TempArr(1 To 1, 1 To 1)
    TempArr(1, 1) = "ERROR, DayHourMin must end with D, H or M"
    C_ARR_OHLCV = TempArr
    Exit Function
End If

AggrTxt = ""
If Len(DayHourMin) > 1 Then
    AggrVal = Left(DayHourMin, Len(DayHourMin) - 1)
    If Val(AggrVal) >= 1 And Val(AggrVal) <= 60 Then
        AggrTxt = "&aggregate=" & Val(AggrVal)
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
    TimeTxt = "&toTs=" & dt
Else
    TimeTxt = ""
End If

If Len(exchange) > 2 Then
    ExchangeTxt = "&e=" & exchange
Else
    ExchangeTxt = ""
End If
If NrLines > 0 Then
    NrLinesTxt = "&limit=" & NrLines
Else
    NrLinesTxt = ""
End If

PrTxt = PublicCryptoCompareData(cmd, "?fsym=" & CurrBuy & "&tsym=" & CurrSell & AggrTxt & TimeTxt & NrLinesTxt & ExchangeTxt)
'Debug.Print cmd & "?fsym=" & CurrBuy & "&tsym=" & CurrSell & AggrTxt & TimeTxt & NrLinesTxt & ExchangeTxt
'Debug.Print PrTxt
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
                    TempArr(j, i) = ResTbl(itmnr, j)
                    If itmnr = 2 Then
                        'Time from Unixtime to normal date/time
                        If j > 1 Then
                            utime = ResTbl(itmnr + 1, j)
                            TempArr(j, i) = UnixTimeToDate(utime)
                        Else
                            TempArr(j, i) = ResTbl(itmnr + 1, j)
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

