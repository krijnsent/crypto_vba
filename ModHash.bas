Attribute VB_Name = "ModHash"
'Public Function Suite() As TestSuite
'  Set Suite = New TestSuite
'  Suite.Description = "..."
'
'  ' Create reporter and attach it to these specs
'  Dim Reporter As New ImmediateReporter
'  Reporter.ListenTo Suite
'
'
'  ' -> Reporter will now output results as they are generated
'End Function

Sub TestHash()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModHash"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestHashes")

'9f54d278014e50f71c789e6fba09c6cfb0945d9253eb8dc5f91ecf52e9996ab9
TestResult = ComputeHash_C("SHA256", "input_string", "", "STRHEX")
Test.IsEqual Len(TestResult), 64
Test.IsEqual TestResult, "9f54d278014e50f71c789e6fba09c6cfb0945d9253eb8dc5f91ecf52e9996ab9"

'9DsHyKCMZmDa5+y2I4v9ErMAa4rTWXVZVqDA5HOuScHFJBjUJeJW11B6CojHJHQHIzXJc8tkneRLRCqaZfV05A==
TestResult = ComputeHash_C("SHA512", "input_string", "my_key", "STR64")
Test.IsEqual Len(TestResult), 88
Test.IsEqual TestResult, "9DsHyKCMZmDa5+y2I4v9ErMAa4rTWXVZVqDA5HOuScHFJBjUJeJW11B6CojHJHQHIzXJc8tkneRLRCqaZfV05A=="

'2•9uêDÍ{S®—¢9ôK˙À≠ìS’©Üåk¨46gë°yR˛Êâe∂ÆÚû˙ﬂ
TestResult = ComputeHash_C("SHA384", "input_string", "", "RAW")
'If Len(TestResult) = 48 And Left(TestResult, 4) = "2•9u" Then
Test.IsEqual Len(TestResult), 48
Test.IsEqual Left(TestResult, 4), "2•9u"

End Sub

Function ComputeHash_C(Meth As String, ByVal clearText As String, ByVal key As String, Optional OutType As String) As Variant

    'Created by Koen Rijnsent, www.castoro.nl
    'Function to return a hash
    'Methods: default SHA1, other options SHA512, SHA384 and SHA256
    'Key: "" for no key
    'Output: STR64, STRHEX, RAW or bytes
    
    Dim BKey() As Byte
    Dim BTxt() As Byte
    
    Dim oT As Object
    Dim TextToHash() As Byte
    Dim bytes() As Byte
    
    BTxt = StrConv(clearText, vbFromUnicode)
    BKey = StrConv(key, vbFromUnicode)
    
    If key <> "" Then
        'MD5 does not work with a key, no error catching yet
        If Meth = "SHA512" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.HMACSHA512")
        ElseIf Meth = "SHA384" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.HMACSHA384")
        ElseIf Meth = "SHA256" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.HMACSHA256")
        Else
            Set SHAhasher = CreateObject("System.Security.Cryptography.HMACSHA1")
        End If
        SHAhasher.key = BKey
        bytes = SHAhasher.computeHash_2(BTxt)
    Else
        If Meth = "SHA512" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.SHA512Managed")
        ElseIf Meth = "SHA256" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.SHA256Managed")
        ElseIf Meth = "SHA384" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.SHA384Managed")
        ElseIf Meth = "MD5" Then
            Set SHAhasher = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
        Else
            Set SHAhasher = CreateObject("System.Security.Cryptography.SHA1Managed")
        End If
        Set oT = CreateObject("System.Text.UTF8Encoding")
        TextToHash = oT.GetBytes_4(clearText)
        bytes = SHAhasher.computeHash_2((TextToHash))
    End If
    
    If OutType = "STR64" Then
       ComputeHash_C = ConvToBase64String(bytes)
    ElseIf OutType = "STRHEX" Then
       ComputeHash_C = ConvToHexString(bytes)
    ElseIf OutType = "RAW" Then
        ComputeHash_C = Base64Decode(ConvToBase64String(bytes))
    Else
       ComputeHash_C = bytes
    End If
    Set SHAhaser = Nothing

End Function

Function ConvToBase64String(vIn As Variant) As Variant

    'Source: https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/File_Hashing_in_VBA
    Dim oD As Object
    Set oD = CreateObject("MSXML2.DOMDocument")
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.base64"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToBase64String = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function

Function ConvToHexString(vIn As Variant) As Variant

    'Source: https://en.wikibooks.org/wiki/Visual_Basic_for_Applications/File_Hashing_in_VBA
    Dim oD As Object
    Set oD = CreateObject("MSXML2.DOMDocument")
      
      With oD
        .LoadXML "<root />"
        .DocumentElement.DataType = "bin.Hex"
        .DocumentElement.nodeTypedValue = vIn
      End With
    ConvToHexString = Replace(oD.DocumentElement.Text, vbLf, "")
    
    Set oD = Nothing

End Function


' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Function Base64Encode(inData)
  'rfc1521
  '2001 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim cOut, sOut, i
  
  'For each group of 3 bytes
  For i = 1 To Len(inData) Step 3
    Dim nGroup, pOut, sGroup
    
    'Create one long from this 3 bytes.
    nGroup = &H10000 * Asc(Mid(inData, i, 1)) + _
      &H100 * MyASC(Mid(inData, i + 1, 1)) + MyASC(Mid(inData, i + 2, 1))
    
    'Oct splits the long To 8 groups with 3 bits
    nGroup = Oct(nGroup)
    
    'Add leading zeros
    nGroup = String(8 - Len(nGroup), "0") & nGroup
    
    'Convert To base64
    pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _
      Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1)
    
    'Add the part To OutPut string
    sOut = sOut + pOut
    
    'Add a new line For Each 76 chars In dest (76*3/4 = 57)
    'If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf
  Next
  Select Case Len(inData) Mod 3
    Case 1: '8 bit final
      sOut = Left(sOut, Len(sOut) - 2) + "=="
    Case 2: '16 bit final
      sOut = Left(sOut, Len(sOut) - 1) + "="
  End Select
  Base64Encode = sOut
End Function

Function MyASC(OneChar)
  If OneChar = "" Then MyASC = 0 Else MyASC = Asc(OneChar)
End Function


