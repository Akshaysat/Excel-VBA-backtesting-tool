' use this to save downloaded data to csv
Sub WriteStringToTextFile(s As String, fileName As String, Optional folderName As String)
    
    Dim myFile As String
    
    ' saves in documents, subfolder optional
    If Not IsMissing(folderName) Then
        myFile = Application.DefaultFilePath & "\" & folderName & "\" & fileName
    Else
        myFile = Application.DefaultFilePath & fileName
    End If
    
    Open myFile For Output As #1
    
    Write #1, s
    
    Close #1
    
End Sub
