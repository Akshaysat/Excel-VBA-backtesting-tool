' a collection of trades
' and all the analytics on them
Private pTrades As Collection

Public Sub InitiateProperties()
    Set pTrades = New Collection
    
End Sub

Public Function NumberOfTrades() As Long
    NumberOfTrades = pTrades.Count
End Function

Public Function Position() As Long
    If pTrades.Count = 0 Then
        Position = 0
    Else
        Dim t As Trade
        For Each t In pTrades
            Position = Position + t.Position
        Next t
    End If
End Function

Public Sub AddTrade(MyTrade As Trade)
    pTrades.Add MyTrade
End Sub
