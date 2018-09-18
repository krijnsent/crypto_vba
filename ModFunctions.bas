Attribute VB_Name = "ModFunctions"
'Functions in module:
'DateToUnixTime - retuns the UnixTime of a date/time
'UnixTimeToDate - returns the date/time of a UnixTime
'TransposeArr - Custom transpose function, worksheetfunction.transpose won't handle long strings
'URLEncode - especially for Excel 2013 and before, afterwards it's a standard function
'Source: https://github.com/krijnsent/crypto_vba
Sub TestFunctions()

Debug.Print "TestFunctions"

TestResult = CreateNonce()
'151802369827
If Len(TestResult) = 12 And TestResult > 151802369827# Then
    Debug.Print "OK"
Else
    Debug.Print "ERROR"
End If

TestResult = CreateNonce("10")
'1518023698
If Len(TestResult) = 10 And TestResult > 1518023698 Then
    Debug.Print "OK"
Else
    Debug.Print "ERROR"
End If

TestResult = CreateNonce(3)
'151
If Len(TestResult) = 3 And TestResult >= 151 Then
    Debug.Print "OK"
Else
    Debug.Print "ERROR"
End If

TestResult = CreateNonce(15)
'151802369828000
If Len(TestResult) = 15 And TestResult >= 151802369828000# Then
    Debug.Print "OK"
Else
    Debug.Print "ERROR"
End If

Debug.Print DateToUnixTime(#4/26/2017#)
'1493164800
Debug.Print DateToUnixTime(Now)
'e.g. 1511958343
Debug.Print UnixTimeToDate(1493164800)
'26-4-2017
Debug.Print UnixTimeToDate(1511958343)
'29-11-2017 12:25:43

' Declare a two dimensional array
' Fill the array with text made up of i and j values
Dim TestArr(1 To 3, 1 To 2) As Variant
Dim i As Long, j As Long
For i = LBound(TestArr) To UBound(TestArr)
    For j = LBound(TestArr, 2) To UBound(TestArr, 2)
        TestArr(i, j) = CStr(i) & ":" & CStr(j)
    Next j
Next i
FlipArr = TransposeArr(TestArr)
Debug.Print TestArr(1, 2)
Debug.Print FlipArr(2, 1)

Debug.Print URLEncode("http://www.github.com/")
'http%3A%2F%2Fwww.github.com%2F
Debug.Print URLEncode("https://github.com/search?q=crypto_vba&type=")
'https%3A%2F%2Fgithub.com%2Fsearch%3Fq%3Dcrypto_vba%26type%3D

End Sub

Function DateToUnixTime(dt) As Long
    DateToUnixTime = 0
    On Error Resume Next
    DateToUnixTime = DateDiff("s", "1/1/1970", dt)
    On Error GoTo 0
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

Function CreateNonce(Optional NonceLength As Integer = 12) As String
    
    Dim ScsLng As Long
    ScsLng = Int(Timer() * 100)
    
    NonceUnique = DateDiff("s", "1/1/1970", Now)
    If NonceLength >= 12 Then
        CreateNonce = NonceUnique & Right(ScsLng, 2) & String(NonceLength - 12, "0")
    ElseIf NonceLength >= 1 Then
        CreateNonce = Left(NonceUnique & Right(ScsLng, 2), NonceLength)
    Else
        CreateNonce = 0
    End If

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

