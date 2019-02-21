Attribute VB_Name = "ModWeb"
    'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Based on http://www.808.dk/?code-simplewinhttprequest

Sub TestWeb()

' Create a new test suite
Dim Suite As New TestSuite
Suite.Description = "ModWeb"

' Create reporter and attach it to these specs
Dim Reporter As New ImmediateReporter
Reporter.ListenTo Suite
  
' Create a new test
Dim Test As TestCase
Set Test = Suite.Test("TestWebRequestURL")

'Testing error catching and replies
TestResult = WebRequestURL("myURL", "myMethod")
'{"error_nr":27,"error_txt":"invalid method for WebRequestURL"}
Test.IsEqual Len(TestResult), 62
Test.IsEqual TestResult, "{""error_nr"":27,""error_txt"":""invalid method for WebRequestURL""}"


TestResult = WebRequestURL("myURL", "GET")
'{"error_nr":-2147012796,"error_txt":"VBA-WinHttp.WinHttpRequest  etc.
Test.IsEqual Left(TestResult, 36), "{""error_nr"":-2147012796,""error_txt"":"


TestResult = WebRequestURL("https://github.com/empty_url_not_there", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found"}
Test.IsEqual Len(TestResult), 45
Test.IsEqual TestResult, "{""error_nr"":404,""error_txt"":""HTTP-Not Found""}"


TestResult = WebRequestURL("https://api.kraken.com/0/public/Time", "GET")
'{"error":[],"result":{"unixtime":1511954132,"rfc1123":"Wed, 29 Nov 17 11:15:32 +0000"}}
Test.IsEqual Len(TestResult), 87
Test.IsEqual Left(TestResult, 21), "{""error"":[],""result"":"


Dim headerDict As New Dictionary
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "Customheader", "MyCustomHeader"
TestResult = WebRequestURL("https://httpbin.org/get", "GET", headerDict)
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("url"), "https://httpbin.org/get"
Test.IsEqual JsonResult("headers").Count, 5
Test.IsEqual JsonResult("headers")("Content-Type"), "application/x-www-form-urlencoded"
Test.IsEqual JsonResult("headers")("Customheader"), "MyCustomHeader"


'TEST POST
TestResult = WebRequestURL("https://httpbin.org/post", "POST")
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("url"), "https://httpbin.org/post"
Test.IsEqual JsonResult("headers").Count, 4

Set headerDict = Nothing
headerDict.Add "Content-Type", "application/x-www-form-urlencoded"
headerDict.Add "Customheader", "MyCustomHeader"
TestResult = WebRequestURL("https://httpbin.org/post", "POST", headerDict)
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("url"), "https://httpbin.org/post"
Test.IsEqual JsonResult("headers").Count, 6
Test.IsEqual JsonResult("headers")("Content-Type"), "application/x-www-form-urlencoded"
Test.IsEqual JsonResult("headers")("Customheader"), "MyCustomHeader"


TestResult = WebRequestURL("https://httpbin.org/post", "POST", , "my_post_message")
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("url"), "https://httpbin.org/post"
Test.IsEqual JsonResult("data"), "my_post_message"
Test.IsEqual JsonResult("headers").Count, 5


TestResult = WebRequestURL("https://httpbin.org/post", "POST", headerDict, "my_post_message_2=msg")
Set JsonResult = JsonConverter.ParseJson(TestResult)
Test.IsEqual JsonResult("url"), "https://httpbin.org/post"
Test.IsEqual JsonResult("form")("my_post_message_2"), "msg"
Test.IsEqual JsonResult("headers").Count, 6
Test.IsEqual JsonResult("headers")("Customheader"), "MyCustomHeader"


End Sub


Function WebRequestURL(strURL As String, strMethod As String, Optional objHeaders As Dictionary, Optional strPostMsg As String) As String

' Instantiate a WinHttpRequest object and open it
ErrResp = "{""error_nr"":ERR_NR,""error_txt"":""ERR_TXT""}"
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
If strMethod = "GET" Then
    On Error Resume Next
    objHTTP.Open "GET", strURL
    If Not objHeaders Is Nothing Then
        For Each Key In objHeaders.Keys()
            'e.g. objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            objHTTP.setRequestHeader Key, objHeaders(Key)
        Next Key
    Else
        'No headers
    End If
    
    objHTTP.Send
    If Err.Number = 0 Then
        If objHTTP.Status = "200" Then
            objHTTP.WaitForResponse
            WebRequestURL = objHTTP.ResponseText
        Else
            WebRequestURL = Replace(Replace(ErrResp, "ERR_NR", objHTTP.Status), "ERR_TXT", "HTTP-" & objHTTP.StatusText)
        End If
    Else
        'Unknown error, probably no internet connection, answer in JSON
        WebRequestURL = Replace(Replace(ErrResp, "ERR_NR", Err.Number), "ERR_TXT", "VBA-" & Err.Source & " " & Err.Description)
    End If
    On Error GoTo 0
ElseIf strMethod = "POST" Then
    On Error Resume Next
    objHTTP.Open "POST", strURL
    If Not objHeaders Is Nothing Then
        For Each Key In objHeaders.Keys()
            'e.g. objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
            objHTTP.setRequestHeader Key, objHeaders(Key)
        Next Key
    Else
        'No headers
    End If
    
    If strPostMsg = "" Then
        objHTTP.Send
    Else
        objHTTP.Send (strPostMsg)
    End If

    If Err.Number = 0 Then
        If objHTTP.Status = "200" Then
            objHTTP.WaitForResponse
            WebRequestURL = objHTTP.ResponseText
        Else
            WebRequestURL = Replace(Replace(ErrResp, "ERR_NR", objHTTP.Status), "ERR_TXT", "HTTP-" & objHTTP.StatusText)
        End If
    Else
        'Unknown error, probably no internet connection, answer in JSON
        WebRequestURL = Replace(Replace(ErrResp, "ERR_NR", Err.Number), "ERR_TXT", "VBA-" & Err.Source & " " & Err.Description)
    End If
    On Error GoTo 0

'   Not implemented, work in progress
'    objHTTP.Open "POST", Url, False
'    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
'    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    objHTTP.setRequestHeader "Key", apikey
'    objHTTP.setRequestHeader "Sign", APIsign
'    objHTTP.Send (postdata)
Else
    WebRequestURL = Replace(Replace(ErrResp, "ERR_NR", 27), "ERR_TXT", "invalid method for WebRequestURL")
End If
Set objHTTP = Nothing

End Function

