Attribute VB_Name = "ModFunctions"
'Functions in module:
'DateToUnixTime - retuns the UnixTime of a date/time
'UnixTimeToDate - returns the date/time of a UnixTime
'TransposeArr - Custom transpose function, worksheetfunction.transpose won't handle long strings
'URLEncode - especially for Excel 2013 and before, afterwards it's a standard function
'Source: https://github.com/krijnsent/crypto_vba
Sub TestFunctions()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModFunctions"
' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
' Create a new test
Dim Test As TestCase


Set Test = Suite.Test("CreateNonce")
TestResult = CreateNonce()
Test.IsOk TestResult > 151802369827#
Test.IsEqual Len(TestResult), 12

TestResult = CreateNonce("10")
Test.IsOk TestResult > 1518023698
Test.IsEqual Len(TestResult), 10

TestResult = CreateNonce(3)
Test.IsOk TestResult >= 151
Test.IsEqual Len(TestResult), 3

TestResult = CreateNonce(15)
Test.IsOk TestResult > 151802369827000#
Test.IsEqual Len(TestResult), 15


Set Test = Suite.Test("DateToUnixTime")
TestResult = DateToUnixTime(#4/26/2017#)
Test.IsEqual TestResult, 1493164800
Test.IsEqual Len(TestResult), 10

TestResult = DateToUnixTime(Now)
Test.IsOk TestResult > 1511958343
Test.IsEqual Len(TestResult), 10


Set Test = Suite.Test("UnixTimeToDate")
TestResult = UnixTimeToDate(1493164800)
Test.IsEqual TestResult, #4/26/2017#
Test.IsEqual Len(TestResult), 9

TestResult = UnixTimeToDate(1511958343)
Test.IsEqual TestResult, #11/29/2017 12:25:43 PM#
Test.IsEqual Len(TestResult), 19


Set Test = Suite.Test("TransposeArr")
' Declare a two dimensional array, Fill the array with text made up of i and j values
Dim TestArr(1 To 3, 1 To 2) As Variant
Dim i As Long, j As Long
For i = LBound(TestArr) To UBound(TestArr)
    For j = LBound(TestArr, 2) To UBound(TestArr, 2)
        TestArr(i, j) = CStr(i) & ":" & CStr(j)
    Next j
Next i
FlipArr = TransposeArr(TestArr)
Test.IsEqual TestArr(1, 2), "1:2"
Test.IsEqual TestArr(1, 2), FlipArr(2, 1)

'Test URLEncode
Set Test = Suite.Test("URLEncode")
TestResult = URLEncode("http://www.github.com/")
Test.IsEqual TestResult, "http%3A%2F%2Fwww.github.com%2F"

TestResult = URLEncode("https://github.com/search?q=crypto_vba&type=")
Test.IsEqual TestResult, "https%3A%2F%2Fgithub.com%2Fsearch%3Fq%3Dcrypto_vba%26type%3D"


'TestDictToString
Set Test = Suite.Test("TestDictToString")
Dim testDict As New Dictionary

'Empty Dict
TestResult = DictToString(testDict, "JSON")
Test.IsEqual TestResult, "{}"

'Unknown type
TestResult = DictToString(testDict, "-")
Test.IsEqual TestResult, "UNKNOWN_TYPE"

'Fill dictionary
testDict.Add "option1", "BTC-ETH"
testDict.Add "another_option", "16"
JsonTxt = "{""option1"":""BTC-ETH"",""another_option"":""16""}"
TestResult = DictToString(testDict, "JSON")
Test.IsEqual TestResult, JsonTxt
UrlTxt = "option1=BTC-ETH&another_option=16"
TestResult = DictToString(testDict, "URLENC")
Test.IsEqual TestResult, UrlTxt

Dim testDict2 As New Dictionary
'Fill dictionary
testDict2.Add "value1", 9
testDict2.Add "value_2", 0.154
testDict2.Add "value_as_string", "1.87"
testDict2.Add "commaval_as_str", "2,16"
TestResult = DictToString(testDict2, "JSON")
JsonTxt = "{""value1"":9,""value_2"":0.154,""value_as_string"":""1.87"",""commaval_as_str"":""2,16""}"
Test.IsEqual TestResult, JsonTxt
TestResult = DictToString(testDict2, "URLENC")
UrlTxt = "value1=9&value_2=0.154&value_as_string=1.87&commaval_as_str=2,16"
Test.IsEqual TestResult, UrlTxt


'TestSortDict
Set Test = Suite.Test("TestSortDict")

'Function: Sort dictionaries
Dim testDict3 As New Dictionary
'Fill dictionary
testDict3.Add "d", 9
testDict3.Add "e", 0.154
testDict3.Add "c", "1.87"
testDict3.Add "b", "2,16"

'Sort normally
Call SortDictByKey(testDict3)
Test.IsEqual testDict3.Count, 4
Test.IsEqual testDict3.Keys(0), "b"
Test.IsEqual testDict3.Keys(3), "e"
Test.IsEqual testDict3.Items(3), 0.154

'Sort desc
Call SortDictByKey(testDict3, False)
Test.IsEqual testDict3.Count, 4
Test.IsEqual testDict3.Keys(0), "e"
Test.IsEqual testDict3.Keys(3), "b"
Test.IsEqual testDict3.Items(3), "2,16"


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


Function DictToString(DictIn As Dictionary, OutputType As String) As String

Dim OutputTxt As String
Dim ValStr As String

If DictIn Is Nothing Then
    DictToString = ""
    Exit Function
End If

If OutputType = "JSON" Then
    OutputTxt = "{"
    For Each opt In DictIn.Keys()
        If OutputTxt <> "{" Then OutputTxt = OutputTxt & ","
        'If a string came in, put double quotes around it
        ValD = DictIn(opt)
        Separ = ""
        If VarType(ValD) = vbString Then Separ = """"
        
        'Value: correct for comma decimal system if a value was supplied
        ValStr = ValD
        If VarType(ValD) <> vbString Then ValStr = Replace(ValStr, ",", ".")
        OutputTxt = OutputTxt & """" & opt & """" & ":" & Separ & ValStr & Separ
    Next
    OutputTxt = OutputTxt & "}"
ElseIf OutputType = "URLENC" Then
    OutputTxt = ""
    For Each opt In DictIn.Keys()
        If OutputTxt <> "" Then OutputTxt = OutputTxt & "&"
        ValD = DictIn(opt)
        ValStr = ValD
        If VarType(ValD) <> vbString Then ValStr = Replace(ValStr, ",", ".")
        OutputTxt = OutputTxt & opt & "=" & ValStr
    Next
Else
    OutputTxt = "UNKNOWN_TYPE"
End If

DictToString = OutputTxt

End Function


Sub SortDictByKey(DictIn As Dictionary, Optional SortAsc As Boolean = True)
    'Default: sort dictionary Ascending by Key
    'Inspired by https://excelmacromastery.com/vba-dictionary/#Sorting_the_Dictionary
    
    Dim ResDict As New Dictionary
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    'Exit if DictIn is empty or only has max 1 item
    If DictIn Is Nothing Then
        Exit Sub
    Else
        If DictIn.Count <= 1 Then
            Exit Sub
        End If
    End If
    
    ' Put keys in array and sort (asc/desc)
    For Each Key In DictIn.Keys
        arrayList.Add Key
    Next Key
    arrayList.Sort
    If SortAsc = False Then
        arrayList.Reverse
    End If
    
    'Loop through array
    For Each va In arrayList
        ResDict.Add va, DictIn(va)
    Next va
    
    Set DictIn = ResDict

End Sub

