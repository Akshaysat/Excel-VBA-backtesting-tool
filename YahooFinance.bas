Option Explicit
' call yahoo finance to get the data

' gets all available data on a particular symbol
Function YahooFinanceRequest(symbol As String) As String
 
    Dim url As String
    Dim m As Integer, d As Integer, y As Integer
    Dim dd As Date
    dd = Now
    
 ' month is 0 based
    m = Month(dd) - 1
    d = Day(dd)
    y = Year(dd)
    
    url = "http://ichart.finance.yahoo.com/table.csv?s={symb}&a=1&b=1&c=1900&d={month}&e={day}&f={year}&g=d&ignore=.csv"

    url = Replace(url, "{symb}", symbol)
    url = Replace(url, "{month}", CStr(m))
    url = Replace(url, "{day}", CStr(d))
    url = Replace(url, "{year}", CStr(y))
    
    Dim Http As Object
    Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Http.Open "GET", url, False
    Http.send
    YahooFinanceRequest = Http.responsetext

End Function
