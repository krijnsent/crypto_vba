Attribute VB_Name = "ModExchCoinigy"
Sub TestCoinigy()

'Source: https://github.com/krijnsent/crypto_vba
'Remember to create a new API key for excel/VBA

Dim apiKey As String
Dim secretKey As String

apiKey = "your api key here"
secretKey = "your secret key here"

'Remove these 2 lines, unless you define 2 constants somewhere ( Public Const secretkey_btce = "the key to use everywhere" etc )
apiKey = apikey_coinigy
secretKey = secretkey_coinigy

Debug.Print PublicCoinigy("exchanges", "")
'{"error_nr":999,"error_txt":"PublicCoinigy does not exist on Coinigy"}

'Unix time period:
t1 = DateToUnixTime("1/1/2014")
t2 = DateToUnixTime("1/1/2018")

Debug.Print PrivateCoinigy("exchanges", apiKey, secretKey)
'{"data":[{"exch_id":"4","exch_name":"Bitstamp","exch_code":"BITS","exch_fee":"0.0025","exch_trade_enabled":"1","exch_balance_enabled":"1","exch_url":"https:\/\/www.bitstamp.net\/"},{"exch_id":"7","exch_name":"Bitfinex","exch_code":"BITF","exch_fee":"0.003","exch_trade_enabled":"1","exch_balance_enabled":"1","exch_url":"https:\/\/www.bitfinex.com\/"},{"exch_id":"11","exch_name":"Kraken","exch_code":"KRKN","exch_fee":"0.003","exch_trade_enabled":"1","exch_balance_enabled":"1","exch_url":"https:\/\/www.kraken.com\/"},{"exch_id":"13","exch_name":"Poloniex","exch_code":"PLNX","exch_fee":"0.002","exch_trade_enabled":"1","exch_balance_enabled":"1","exch_url":"https:\/\/www.poloniex.com\/"},{"exch_id":"15","exch_name":"Bittrex","exch_code":"BTRX","exch_fee":"0.0025","exch_trade_enabled":"1","exch_balance_enabled":"1","exch_url":"https:\/\/bittrex.com\/"},{"exch_id":"16","exch_name":"C-Cex","exch_code":"CCEX",etc...
Debug.Print PrivateCoinigy("markets", apiKey, secretKey, "{""exchange_code"":""BINA""}")
'{"data":[{"exch_id":"62","exch_name":"Global Digital Asset Exchange","exch_code":"GDAX","mkt_id":"720", etc...

End Sub

Function PublicCoinigy(Method As String, Optional MethodOptions As String) As String
'Put here for consitency, does not exist

PublicCoinigy = "{""error_nr"":ERR_NR,""error_txt"":""ERR_TXT""}"
PublicCoinigy = Replace(Replace(PublicCoinigy, "ERR_NR", 999), "ERR_TXT", "PublicCoinigy does not exist on Coinigy")

End Function
Function PrivateCoinigy(Method As String, apiKey As String, secretKey As String, Optional MethodOptions As String) As String

Dim NonceUnique As String
'https://api.coinigy.com/api/v1/

TradeApiSite = "https://api.coinigy.com/api/v1/"
Url = TradeApiSite & Method
postdata = MethodOptions

' Instantiate a WinHttpRequest object and open it
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
'Debug.Print "POST: " & TradeApiSite & postdata
objHTTP.Open "POST", Url, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-Type", "application/json"
objHTTP.setRequestHeader "X-API-KEY", apiKey
objHTTP.setRequestHeader "X-API-SECRET", secretKey
objHTTP.Send (postdata)

objHTTP.WaitForResponse
PrivateCoinigy = objHTTP.ResponseText
Set objHTTP = Nothing

End Function


