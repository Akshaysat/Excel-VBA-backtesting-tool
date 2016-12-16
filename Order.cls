Option Explicit


' quantity - positive means buy, negative means sell
Private pQuantity As Long
' symbol
Private pSymbol As String
' price
Private pPrice As Double
' executions
Private pDate As Date

Private pExecutions As Collection

Public Property Get OrderQuantity() As Long
    OrderQuantity = pQuantity
End Property

Public Property Get OrderPrice() As Double
    OrderPrice = pPrice
End Property

Public Property Get OrderSymbol() As String
    OrderSymbol = pSymbol
End Property

Public Property Let OrderDate(OrderDate As Date)
    pDate = OrderDate
End Property

Public Property Get OrderDate() As Date
    OrderDate = pDate
End Property


Public Sub InitiateProperties( _
    Quantity As Long, _
    Symbol As String, _
    Price As Double, _
    OrderDate As Date)
    
    pQuantity = Quantity
    pSymbol = Symbol
    pPrice = Price
    Date = OrderDate
    
    Set pExecutions = New Collection
    
End Sub

Public Sub AddExecution(e As Execution)
    pExecutions.Add e
End Sub

Public Function ExecutionQuantity() As Long
    Dim e As Execution
    For Each e In pExecutions
        ExecutionQuantity = e.Quantity
    Next e
End Function

Public Function DollarValue() As Double
    Dim e As Execution
    For Each e In pExecutions
        DollarValue = DollarValue + e.DollarValue
    Next e
End Function

Public Function FirstExecutionDate() As Date
    If pExecutions.Count > 0 Then
        Dim e As Execution
        Set e = pExecutions(1)
        FirstExecutionDate = e.TheDate
    End If
End Function
    
Public Function IsFilled() As Boolean
    If pQuantity > 0 And ExecutionQuantity = pQuantity Then IsFilled = True
End Function

