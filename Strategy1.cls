Option Explicit

Private pDataCollection As Collection
' position is what? position is long, short or flat
' but a position is more than that . . .
' that's a trade
Private pToday As EndOfDayData
Private pMA20Close As MovingAverage
Private pADR As MovingAverage
Private pSymbol As String
' orders are connected to trades
' Private pOrders As Collection
' need a collection of trades/trades are what will be analyzed
Private pOpenTrade As Trade
Private pClosedTrades As Trades

Private Const DollarsPerPosition = 10000

' initialize the collection -
' collection is passed into initializer
Public Sub InitiateProperties(DataCollection As Collection, Symbol As String)

    Set pDataCollection = DataCollection
    Set pMA20Close = Factory.CreateMovingAverage(20)
    Set pADR = Factory.CreateMovingAverage(20)
    pSymbol = Symbol
    Set pClosedTrades = Factory.CreateTrades
    
End Sub

Public Sub Run()
    Dim eodd As EndOfDayData
    For Each eodd In pDataCollection
        Debug.Print eodd.DateOfData
        OnData eodd
        
    Next eodd
End Sub

Public Sub OnData(eodd As EndOfDayData)
    
    Dim OrderPrice As Double
    Dim OrderShares As Long
    Dim o As Order
    Dim e As Execution
        
    Set pToday = eodd
    
    ' update indicators
    pMA20Close.AddDataPoint eodd.ClosePrice
    pADR.AddDataPoint (eodd.HighPrice - eodd.LowPrice)
    
    ' check entry conditions
    If EntryConditions = True Then
         
        ' create order
        
        ' calculate order share and price
        OrderPrice = Round(pMA20Close.Value - pADR.Value, 2)
        OrderShares = Round(DollarsPerPosition / OrderPrice, 0)
         
        Set o = Factory.CreateOrder(OrderShares, pSymbol, OrderPrice, eodd.DateOfData)
    
        Dim t As Trade
        Set t = Factory.CreateTrade(o)
        
        Set pOpenTrade = t
        
        ' force immediate fill, send a mock execution
 
        Set e = Factory.CreateExecution(o.OrderQuantity, o.OrderSymbol, o.OrderPrice, pToday.DateOfData)
        o.AddExecution e
    End If
    
    ' check exit conditions
    ' is not flat
    ' if the same date, close > executionprice + adr
    ' if another date, high > execution price + adr
    If ExitConditions Then
               
        ' create order
        
        ' calculate order share and price
        ' order price is the maximum of limit price or where the stock opens at
        If pOpenTrade.FirstEntryDate = eodd.DateOfData Then
            OrderPrice = pOpenTrade.AverageEntryPrice + pADR.Value
        Else
            OrderPrice = MinMax.Max( _
                eodd.OpenPrice, _
                pOpenTrade.AverageEntryPrice + pADR.Value)
        End If
        OrderShares = Round(pOpenTrade.Position, 0)
        
        
        Set o = Factory.CreateOrder(OrderShares, Me.Symbol, OrderPrice, eodd.DateOfData)
        
        pOpenTrade.AddOrder o
        
        ' force immediate fill, send a mock execution
        
        Set e = Factory.CreateExecution(OrderShares, Me.Symbol, OrderPrice, pToday.DateOfData)
        
        o.AddExecution e
        pOpenTrade.AddOrder o
        
        If pOpenTrade.IsClosed Then
            pClosedTrades.AddTrade pOpenTrade
            Set pOpenTrade = Nothing
        End If
        
    End If
              
     
End Sub

Public Function EntryConditions() As Boolean

    If Not IsFlat Then Exit Function
    
    If Not pMA20Close.IsInitialized Then Exit Function
    
    If Not pADR.IsInitialized Then Exit Function
    
    If Not pToday.LowPrice < pMA20Close.Value - pADR.Value Then Exit Function
    
    EntryConditions = True

End Function

Public Function ExitConditions() As Boolean

    If IsFlat Then Exit Function
    
    
     
    Dim EntryDate As Date
    EntryDate = pOpenTrade.FirstEntryDate
    
    Dim EntryPrice As Double
    EntryPrice = pOpenTrade.AverageEntryPrice
    
    Dim TodaysDate As Date
    TodaysDate = pToday.DateOfData
    
    If ( _
        ( _
            EntryDate = TodaysDate _
            And _
            pToday.ClosePrice > EntryPrice + pADR.Value _
        ) _
            Or _
        ( _
            EntryDate < pToday.DateOfData _
            And _
            pToday.HighPrice > EntryPrice + pADR.Value _
        ) _
    ) = False _
               Then Exit Function
               
    ExitConditions = True
End Function


Public Property Get Symbol() As Boolean

End Property

' for the purpose of backtesting end of day data this order will always assume fill
Public Sub OrderImmediateFill(Price As Double, Shares As Long)
    
End Sub

 

Private Function MovingAverageInitialized() As Boolean
    
End Function

Private Function IsFlat() As Boolean
    
    If pOpenTrade Is Nothing Then
        IsFlat = True ' because no trade no position so  you're flat
        Exit Function
    End If
    IsFlat = pOpenTrade.Position = 0

End Function
 



