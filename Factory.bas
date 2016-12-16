Option Explicit

Public Function CreateEndOfDayData( _
    DateOfData As Date, _
    OpenPrice As Double, _
    High As Double, _
    Low As Double, _
    ClosePrice As Double, _
    Volume As Long, _
    SplitAdjustedPrice As Double _
) As EndOfDayData

Dim endOfDayDataObject As EndOfDayData
Set endOfDayDataObject = New EndOfDayData

endOfDayDataObject.InitiateProperties _
    DateOfData, _
    OpenPrice, _
    High, _
    Low, _
    ClosePrice, _
    Volume, _
    SplitAdjustedPrice
 
Set CreateEndOfDayData = endOfDayDataObject

End Function

Public Function CreateMovingAverage(Length As Double)

Dim ma As MovingAverage
Set ma = New MovingAverage

ma.InitiateProperties (Length)

Set CreateMovingAverage = ma

End Function

Public Function CreateExecution(Quantity As Long, Symbol As String, Price As Double, TheDate As Date)

Dim e As Execution
Set e = New Execution

e.InstantiateProperties Quantity, Symbol, Price, TheDate

Set CreateExecution = e

End Function

Public Function CreateOrder(Quantity As Long, Symbol As String, Price As Double, OrderDate As Date)

Dim o As Order
Set o = New Order

o.InitiateProperties Quantity, Symbol, Price, OrderDate

Set CreateOrder = o

End Function

Public Function CreateStrategy1(DataCollection As Collection, Symbol As String)

Dim s As Strategy1
Set s = New Strategy1

s.InitiateProperties DataCollection, Symbol

Set CreateStrategy1 = s

End Function

Public Function CreateTrade(MyOrder As Order)

Dim t As Trade
Set t = New Trade

t.InitiateProperties MyOrder

Set CreateTrade = t

End Function

Public Function CreateTrades()

Dim t As Trades
Set t = New Trades

t.InitiateProperties

Set CreateTrades = t

End Function
