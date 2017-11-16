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
    For i = 1 To UBound(ArrIn, 2)
        For j = 1 To UBound(ArrIn, 1)
            TempArr(i, j) = ArrIn(j, i)
        Next
    Next
    
    TransposeArr = TempArr
    
End Function

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
'https://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

