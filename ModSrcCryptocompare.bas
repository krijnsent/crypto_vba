Attribute VB_Name = "ModSrcCryptocompare"
'For historical prices, include https://www.cryptocompare.com/api/

'Functions to use in a sheet
'Source: https://github.com/krijnsent/crypto_vba
'Note: the functions currently slow down the sheets massively, use max 10 functions per workbook, otherwise your workbook might CRASH
'ToDo: include caching - https://fastexcel.wordpress.com/2012/12/05/writing-efficient-udfs-part-12-getting-used-range-fast-using-application-events-and-a-cache/
'ToDo: better error catching

'Functions:
'C_LAST_PRICE -  price?fsym=BTC&tsyms=USD,EUR&e=Coinbase
'C_HIST_PRICE - pricehistorical?fsym=BTC&tsyms=USD,EUR&e=Coinbase&ts=1452680400
'C_DAY_AVG_PRICE - dayAvg?fsym=BTC&tsym=USD&toTs=1487116800&e=Bitfinex
'C_ARR_OHLCV - histoday?fsym=GBP&tsym=USD&limit=30&aggregate=1&e=CCCAGG

Sub GetCryptoCompareFunctionsTest()

'Test for errors first
Set Json = JsonConverter.ParseJson(PublicCryptoCompareData("unknown_command", ""))
Debug.Print Json("Response") & "--" & Json("Message") & "--" & Json("Path")
'Error----/data/unknown_command
Set Json = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", ""))
Debug.Print Json("Response") & "--" & Json("Message") & "--" & Json("Path")
'Error--fsym param seems to be missing.--
Set Json = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "fsym=BTC"))
Debug.Print Json("Response") & "--" & Json("Message") & "--" & Json("Path")
'Error----/data/histodayfsym=BTC
Set Json = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "?fsym=BTC&tsym=BLABLA"))
Debug.Print Json("Response") & "--" & Json("Message") & "--" & Json("Path")
'Error--There is no data for the toSymbol BLABLA .--
Set Json = JsonConverter.ParseJson(PublicCryptoCompareData("histoday", "?fsym=BTC&tsym=XMR"))
Debug.Print Json("Response") & "--" & Json("Message") & "--" & Json("Path")
'Success----

Debug.Print C_LAST_PRICE("MYCOIN1", "BLABLA")
'ERROR There is no data for the symbol MYCOIN1 .
Debug.Print C_LAST_PRICE("BTC", "BLABLA")
'ERROR There is no data for any of the toSymbols BLABLA .
Debug.Print C_LAST_PRICE("BTC", "EUR", "An_Unknown_Exchange")
'ERROR an_unknown_exchange market does not exist for this coin pair (BTC-EUR)
Debug.Print C_LAST_PRICE("BTC", "EUR")
'e.g. 8092,80
Debug.Print C_LAST_PRICE("BTC", "EUR", "Kraken")
'e.g. 8055,90


Debug.Print C_HIST_PRICE("ETH", "USD", "2018-01-01 20:00")
'1056,75
Debug.Print C_HIST_PRICE("ETH", "USD", #1/1/2018#, "Bittrex")
'1111,2

Debug.Print C_DAY_AVG_PRICE("LTC", "BTC", #1/1/2017#)
'0,00452
Debug.Print C_DAY_AVG_PRICE("LTC", "BTC", #1/1/2017#, "Poloniex")
'0,004513

'Function C_ARR_OHLCV(
'DayHour as String, CurrBuy As String, CurrSell As String, ReturnColumns As String -> ETCHLOFV, Optional NrHours As Long,
'Optional MaxTimeDate As Date, Optional Exchange As String) As Variant()
ResArr = C_ARR_OHLCV("H", "BTC", "EUR", "TECV", 48, #1/1/2018#, "Kraken")
Debug.Print ResArr(1, 1), ResArr(1, 2), ResArr(1, 3), ResArr(1, 4)
'time          time          close         volumeto
Debug.Print ResArr(2, 1), ResArr(2, 2), ResArr(2, 3), ResArr(2, 4)
' 1514592000   30-12-2017     12164,33      10076523,28

ResArr = C_ARR_OHLCV("H", "BTC", "EUR", "EC", 24, , "Kraken")
Debug.Print ResArr(1, 1), ResArr(1, 2)
'time          close
Debug.Print ResArr(2, 1), ResArr(2, 2)
'e.g.  31-1-2018 09:00:00           8117

ResArr = C_ARR_OHLCV("H", "BTC", "EUR", "ABD")
Debug.Print ResArr(1, 1), ResArr(1, 2), ResArr(1, 3)
'unknown ReturnColumn unknown ReturnColumn unknown ReturnColumn

ResArr = C_ARR_OHLCV("H", "2FA", "EUR", "ECV")
Debug.Print ResArr(1, 1)
'ERROR There is no data for the symbol 2FA .

ResArr = C_ARR_OHLCV("H", "ETH", "EUR", "")
Debug.Print ResArr(1, 1)
'ERROR ReturnColumns, use the letters ETCHLOFV

ResArr = C_ARR_OHLCV("H", "XLM", "EUR", "TEOHLCFV", 48, DateSerial(2018, 1, 1), "Kraken")
Debug.Print ResArr(1, 1)
'ERROR, cryptocompare API gave back an empty result, try other settings

End Sub

Function PublicCryptoCompareData(Method As String, Optional MethodOptions As String) As String

'https://www.cryptocompare.com/api/ or https://min-api.cryptocompare.com/
Dim Url As String

PublicApiSite = "https://min-api.cryptocompare.com/data"
urlPath = "/" & Method & MethodOptions
Url = PublicApiSite & urlPath

PublicCryptoCompareData = GetDataFromURL(Url, "GET")

Set objHTTP = Nothing

End Function
Function C_LAST_PRICE(CurrBuy As String, CurrSell As String, Optional Exchange As String)

Dim PrTxt As String
Dim Json As Object
Application.Volatile

If Len(Exchange) > 2 Then
    ExchangeTxt = "&e=" & Exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("price", "?fsym=" & CurrBuy & "&tsyms=" & CurrSell & ExchangeTxt)
Set Json = JsonConverter.ParseJson(PrTxt)

If Json("Response") = "Error" Then
    'Error
    C_LAST_PRICE = "ERROR " & Json("Message")
Else
    C_LAST_PRICE = Json(CurrSell)
End If

Set Json = Nothing

End Function

Function C_HIST_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional Exchange As String)

Dim PrTxt As String
Dim Json As Object
Application.Volatile

dt = DateToUnixTime(DateRates)
If Len(Exchange) > 2 Then
    ExchangeTxt = "&e=" & Exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("pricehistorical", "?fsym=" & CurrBuy & "&tsyms=" & CurrSell & "&ts=" & dt & ExchangeTxt)
Set Json = JsonConverter.ParseJson(PrTxt)

If Json("Response") = "Error" Then
    'Error
    C_HIST_PRICE = "ERROR " & Json("Message")
Else
    C_HIST_PRICE = Json(CurrBuy)(CurrSell)
End If

Set Json = Nothing

End Function

Function C_DAY_AVG_PRICE(CurrBuy As String, CurrSell As String, DateRates As Date, Optional Exchange As String)

Dim PrTxt As String
Dim Json As Object
Application.Volatile

dt = DateToUnixTime(DateRates)
If Len(Exchange) > 2 Then
    ExchangeTxt = "&e=" & Exchange
Else
    ExchangeTxt = ""
End If

PrTxt = PublicCryptoCompareData("dayAvg", "?fsym=" & CurrBuy & "&tsym=" & CurrSell & "&toTs=" & dt & ExchangeTxt)
Set Json = JsonConverter.ParseJson(PrTxt)

If Json("Response") = "Error" Then
    'Error
    C_DAY_AVG_PRICE = "ERROR " & Json("Message")
Else
    C_DAY_AVG_PRICE = Json(CurrSell)
End If

Set Json = Nothing

End Function

Function C_ARR_OHLCV(DayHour As String, CurrBuy As String, CurrSell As String, ReturnColumns As String, Optional NrLines As Long, Optional MaxTimeDate As Date, Optional Exchange As String) As Variant()

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
Dim cmd As String
Dim utime As Long
Dim Json As Object
Dim TempArr As Variant
ColumnOptions = "ETCHLOFV"
Application.Volatile

If UCase(DayHour) = "D" Then
    cmd = "histoday"
ElseIf UCase(DayHour) = "H" Then
    cmd = "histohour"
Else
    'Error
    ReDim TempArr(1 To 1, 1 To 1)
    TempArr(1, 1) = "ERROR, DayHour must be D or H"
    C_ARR_OHLCV = TempArr
    Exit Function
End If

If MaxTimeDate > DateSerial(2000, 1, 1) Then
    dt = DateToUnixTime(MaxTimeDate)
    TimeTxt = "&toTs=" & dt
Else
    TimeTxt = ""
End If

If Len(Exchange) > 2 Then
    ExchangeTxt = "&e=" & Exchange
Else
    ExchangeTxt = ""
End If
If NrLines > 0 Then
    NrLinesTxt = "&limit=" & NrLines
Else
    NrLinesTxt = ""
End If

PrTxt = PublicCryptoCompareData(cmd, "?fsym=" & CurrBuy & "&tsym=" & CurrSell & TimeTxt & NrLinesTxt & ExchangeTxt)
'Debug.Print cmd & "?fsym=" & CurrBuy & "&tsym=" & CurrSell & TimeTxt & NrLinesTxt & ExchangeTxt
'Debug.Print PrTxt
Set Json = JsonConverter.ParseJson(PrTxt)

If Json("Response") = "Error" Then
    'Error
    ReDim TempArr(1 To 1, 1 To 1)
    TempArr(1, 1) = "ERROR " & Json("Message")
    C_ARR_OHLCV = TempArr
Else
    If InStr(PrTxt, """Data"":[]") > 0 Then
        'Empty result from Cryptocompare API, show user error
        ReDim TempArr(1 To 1, 1 To 1)
        TempArr(1, 1) = "ERROR, cryptocompare API gave back an empty result, try other settings"
        C_ARR_OHLCV = TempArr
        Exit Function
    End If
    ResArr = JsonToArray(Json)
    ResTbl = ArrayTable(ResArr, True)
    
    ReturnColumns = UCase(Trim(ReturnColumns))
    'Process all columns
    If Len(ReturnColumns) > 0 Then
        ReDim TempArr(1 To UBound(ResTbl, 2), 1 To Len(ReturnColumns))
        For i = 1 To Len(ReturnColumns)
            itm = Mid(ReturnColumns, i, 1)
            itmnr = InStr(ColumnOptions, itm) + 1
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

Set Json = Nothing

End Function

