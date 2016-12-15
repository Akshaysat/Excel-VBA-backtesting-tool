Option Explicit
' call yahoo finance to get the data
' tokenize the results
' use the get function to get desired value


Function YahooFinanceRequest(symbol As String, dd As Date) As String
' returns a csv file in string format for one day
    Dim url As String
    Dim m As Integer, d As Integer, y As Integer
    m = Month(dd) - 1
    d = Day(dd)
    y = Year(dd)
    
    url = "http://ichart.finance.yahoo.com/table.csv?s={symb}&a={month}&b={day}&c={year}&d={month}&e={day}&f={year}&g=d&ignore=.csv"

    url = Replace(url, "{symb}", symbol)
    url = Replace(url, "{month}", CStr(m))
    url = Replace(url, "{day}", CStr(d))
    url = Replace(url, "{year}", CStr(y))
    
    
    YahooFinanceRequest = HttpRequest(url)

End Function
 

Function HttpRequest(url As String) As String
    Dim Http As Object
    Set Http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Http.Open "GET", url, False
    Http.send

    HttpRequest = Http.responsetext
End Function

Function GetOpen(s() As String) As String
    GetOpen = s(7)
End Function

Function GetClose(s() As String) As String
    GetClose = s(10)
End Function

Function Tokenize(queryResult As String) As String()
    Tokenize = Split(queryResult, ",")
End Function
