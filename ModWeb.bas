Attribute VB_Name = "ModWeb"
'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA
'Based on http://www.808.dk/?code-simplewinhttprequest

Sub TestGetData()

'Testing error catching and replies
Debug.Print GetDataFromURL("myURL", "myMethod")
'{"error_nr":27,"error_txt":"invalid method for GetDataFromURL"}
Debug.Print GetDataFromURL("myURL", "GET")
'{"error_nr":-2147012796,"error_txt":"VBA-WinHttp.WinHttpRequest  etc.
Debug.Print GetDataFromURL("https://github.com/empty_url_not_there", "GET")
'{"error_nr":404,"error_txt":"HTTP-Not Found"}

Debug.Print GetDataFromURL("https://api.kraken.com/0/public/Time", "GET")
'{"error":[],"result":{"unixtime":1511954132,"rfc1123":"Wed, 29 Nov 17 11:15:32 +0000"}}

End Sub

Function GetDataFromURL(strURL As String, strMethod As String, Optional strPostData As String) As String

' Instantiate a WinHttpRequest object and open it
ErrResp = "{""error_nr"":ERR_NR,""error_txt"":""ERR_TXT""}"
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
If strMethod = "GET" Then
    On Error Resume Next
    objHTTP.Open "GET", strURL
    objHTTP.Send
    If Err.Number = 0 Then
        If objHTTP.Status = "200" Then
            objHTTP.WaitForResponse
            GetDataFromURL = objHTTP.ResponseText
        Else
            GetDataFromURL = Replace(Replace(ErrResp, "ERR_NR", objHTTP.Status), "ERR_TXT", "HTTP-" & objHTTP.StatusText)
        End If
    Else
        'Unknown error, probably no internet connection, answer in JSON
        GetDataFromURL = Replace(Replace(ErrResp, "ERR_NR", Err.Number), "ERR_TXT", "VBA-" & Err.Source & " " & Err.Description)
    End If
    On Error GoTo 0
ElseIf strMethod = "POST" Then
'   Not implemented, work in progress
'    objHTTP.Open "POST", Url, False
'    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
'    objHTTP.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'    objHTTP.setRequestHeader "Key", apikey
'    objHTTP.setRequestHeader "Sign", APIsign
'    objHTTP.Send (postdata)
Else
    GetDataFromURL = Replace(Replace(ErrResp, "ERR_NR", 27), "ERR_TXT", "invalid method for GetDataFromURL")
End If
Set objHTTP = Nothing

End Function

Function GetDataFromURL_COPY_PASTED(strURL, strMethod, strPostData)
  Dim lngTimeout
  Dim strUserAgentString
  Dim intSslErrorIgnoreFlags
  Dim blnEnableRedirects
  Dim blnEnableHttpsToHttpRedirects
  Dim strHostOverride
  Dim strLogin
  Dim strPassword
  Dim strResponseText
  Dim objWinHttp
  lngTimeout = 59000
  strUserAgentString = "http_requester/0.1"
  intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
  blnEnableRedirects = True
  blnEnableHttpsToHttpRedirects = True
  strHostOverride = ""
  strLogin = ""
  strPassword = ""
  Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
  objWinHttp.SetTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
  objWinHttp.Open strMethod, strURL
  If strMethod = "POST" Then
    objWinHttp.setRequestHeader "Content-type", _
      "application/x-www-form-urlencoded"
  End If
  If strHostOverride <> "" Then
    objWinHttp.setRequestHeader "Host", strHostOverride
  End If
  objWinHttp.Option(0) = strUserAgentString
  objWinHttp.Option(4) = intSslErrorIgnoreFlags
  objWinHttp.Option(6) = blnEnableRedirects
  objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
  If (strLogin <> "") And (strPassword <> "") Then
    objWinHttp.SetCredentials strLogin, strPassword, 0
  End If
  On Error Resume Next
  objWinHttp.Send (strPostData)
  If Err.Number = 0 Then
    If objWinHttp.Status = "200" Then
      GetDataFromURL = objWinHttp.ResponseText
    Else
      GetDataFromURL = "HTTP " & objWinHttp.Status & " " & _
        objWinHttp.StatusText
    End If
  Else
    GetDataFromURL = "Error " & Err.Number & " " & Err.Source & " " & _
      Err.Description
  End If
  On Error GoTo 0
  Set objWinHttp = Nothing
End Function
