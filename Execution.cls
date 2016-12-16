Option Explicit

' symbol
Private pQuantity As Long
Private pSymbol As String
Private pPrice As Double
Private pDate As Date

Public Sub InstantiateProperties( _
    Quantity As Long, _
    Symbol As String, _
    Price As Double, _
    TheDate _
    )
    
    pQuantity = Quantity
    pSymbol = Symbol
    pPrice = Price
    pDate = TheDate
    
End Sub
    
Public Property Get Quantity() As Long
    Quantity = pQuantity
End Property

Public Property Get Symbol() As String
    Symbol = pSymbol
End Property

Public Property Get Price() As Double
    Price = pPrice
End Property

Public Property Get TheDate() As Date
    TheDate = pDate
End Property


Public Function DollarValue() As Double
    DollarValue = pPrice * pQuantity
End Function
