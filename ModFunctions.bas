Attribute VB_Name = "ModFunctions"
Function DateToUnixTime(dt) As Long
    DateToUnixTime = DateDiff("s", "1/1/1970", dt)
End Function

Function UnixTimeToDate(ts As Long) As Date
    'http://www.vbforums.com/showthread.php?513727-RESOLVED-Convert-Unix-Time-to-Date&p=3168062&viewfull=1#post3168062
    Dim intDays As Integer, intHours As Integer, intMins As Integer, intSecs As Integer
    
    intDays = ts \ 86400
    intHours = (ts Mod 86400) \ 3600
    intMins = (ts Mod 3600) \ 60
    intSecs = ts Mod 60
    
    UnixTimeToDate = DateSerial(1970, 1, intDays + 1) + TimeSerial(intHours, intMins, intSecs)
End Function

Function TransposeArr(ArrIn As Variant)

    'Custom transpose function, worksheetfunction.transpose won't handle long strings
    'It will give error 13, https://stackoverflow.com/questions/23315252/vba-tranpose-type-mismatch-error
    Dim TempArr As Variant

    ReDim TempArr(1 To UBound(ArrIn, 2), 1 To UBound(ArrIn, 1))
    For I = 1 To UBound(ArrIn, 2)
        For j = 1 To UBound(ArrIn, 1)
            TempArr(I, j) = ArrIn(j, I)
        Next
    Next
    
    TransposeArr = TempArr
    
End Function
