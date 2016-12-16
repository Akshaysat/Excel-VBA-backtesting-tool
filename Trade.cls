Option Explicit

' trade is a collection of orders
' grouped as such to designate a trade
' o made x$ on xyz means i bought and sold it and then exited and I'm not flat
Private pEntryOrders As Collection
Private pExitOrders As Collection

Public Sub InitiateProperties(o As Order)
    Set pEntryOrders = New Collection
    pEntryOrders.Add o
    ' this is the initial state of a trade, an order submited
    Set pExitOrders = New Collection
End Sub

Public Function Position() As Long
    Dim o As Order
    For Each o In pEntryOrders
        Position = o.ExecutionQuantity
    Next o
    For Each o In pExitOrders
        Position = o.ExecutionQuantity
    Next o
End Function

Public Sub AddOrder(TheOrder As Order)
    Dim o As Order
    Set o = pEntryOrders(1)
    If o.OrderQuantity > 0 Then ' it's a long position
        If TheOrder.OrderQuantity > 0 Then
            ' order is adding
            pEntryOrders.Add TheOrder
        Else
            pExitOrders.Add TheOrder
        End If
    Else ' it's a short position
        If TheOrder.OrderQuantity < 0 Then
            ' order is adding
            pEntryOrders.Add TheOrder
        Else
            pExitOrders.Add TheOrder
        End If
    End If
End Sub

Public Function AverageEntryPrice() As Double
    Dim o As Order
    Dim SumDollarValues As Double, SumExecutedQuantity As Long
    For Each o In pEntryOrders
        SumDollarValues = SumDollarValues + o.DollarValue
        SumExecutedQuantity = SumExecutedQuantity + o.ExecutionQuantity
    Next o
    AverageEntryPrice = SumDollarValues / SumExecutedQuantity
End Function

Public Function AverageExitPrice() As Double
    Dim o As Order
    Dim SumDollarValues As Double, SumExecutedQuantity As Long
    For Each o In pExitOrders
        SumDollarValues = SumDollarValues + o.DollarValue
        SumExecutedQuantity = SumExecutedQuantity + o.ExecutionQuantity
    Next o
    AverageExitPrice = SumDollarValues / SumExecutedQuantity
End Function

Public Function FirstEntryDate() As Date
    ' assume there's at least one order
    Dim o As Order
    Set o = pEntryOrders(1)
    FirstEntryDate = o.FirstExecutionDate
End Function

Public Function EntryExecutionQuantity() As Long
    Dim o As Order
    For Each o In pEntryOrders
        EntryExecutionQuantity = EntryExecutionQuantity + o.ExecutionQuantity
    Next o
End Function

Public Function ExitExecutionQuantity() As Long
    Dim o As Order
    For Each o In pExitOrders
        ExitExecutionQuantity = ExitExecutionQuantity + o.ExecutionQuantity
    Next o
End Function

Public Function IsClosed() As Boolean
    If EntryExecutionQuantity = ExitExecutionQuantity Then IsClosed = True
End Function
