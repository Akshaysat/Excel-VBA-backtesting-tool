' take saved csv files and loads into a collection of class EndOfDayData
' the collection can then be used to test strategies
Function LoadCsvFile(fileName As String, folderName As String) As Variant


    Dim MyData As String, strData() As String, TmpAr() As String
    
    Dim i As Long, n As Long
    
    Dim myFile As String
    If Not IsMissing(folderName) Then
        myFile = Application.DefaultFilePath & "\" & folderName & "\" & fileName
    Else
        myFile = Application.DefaultFilePath & fileName
    End If

    Open myFile For Binary As #1
    
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    strData() = Split(MyData, vbLf)
     
    Dim c As Collection
    Set c = New Collection
    Dim tempArr() As String

    Dim eodd As EndOfDayData
    For n = 1 To UBound(strData) ' skip first row
        tempArr = Split(strData(n), ",")
        If UBound(tempArr) > 0 Then
            Set eodd = New EndOfDayData
            
            
            eodd.DateOfData = DateSerial(Left(tempArr(0), 4), Mid(tempArr(0), 6, 2), Right(tempArr(0), 2))
            
            eodd.OpenPrice = CDbl(tempArr(1))
            eodd.HighPrice = CDbl(tempArr(2))
            eodd.LowPrice = CDbl(tempArr(3))
            eodd.ClosePrice = CDbl(tempArr(4))
            eodd.Volume = CLng(tempArr(5))
            eodd.SplitAdjustedPrice = CDbl(tempArr(6))
             
            
            c.Add eodd
        End If
    Next n
    
    Set LoadCsvFileAsArray = c
     
    
End Function
