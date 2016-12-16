 
Private pDateOfData As Date
Private pOpen As Double
Private pHigh As Double
Private pLow As Double
Private pClose As Double
Private pVolume As Long
Private pSplitAdjustedPrice As Double
 

Public Property Get DateOfData() As Date
    DateOfData = pDateOfData
End Property
Public Property Let DateOfData(Value As Date)
    pDateOfData = Value
End Property

Public Property Get OpenPrice() As Double
    OpenPrice = pOpen
End Property
Public Property Let OpenPrice(Value As Double)
    pOpen = Value
End Property
Public Property Get HighPrice() As Double
    HighPrice = pHigh
End Property
Public Property Let HighPrice(Value As Double)
    pHigh = Value
End Property
Public Property Get LowPrice() As Double
    LowPrice = pLow
End Property
Public Property Let LowPrice(Value As Double)
    pLow = Value
End Property
Public Property Get ClosePrice() As Double
    ClosePrice = pClose
End Property
Public Property Let ClosePrice(Value As Double)
    pClose = Value
End Property
Public Property Get SplitAdjustedPrice() As Double
    SplitAdjustedPricePrice = pSplitAdjustedPrice
End Property
Public Property Let SplitAdjustedPrice(Value As Double)
    pSplitAdjustedPrice = Value
End Property
Public Property Get Volume() As Long
    Volume = pVolume
End Property
Public Property Let Volume(Value As Long)
    pVolume = Value
End Property

Public Sub InitiateProperties( _
    DateOfData As Date, _
    OpenPrice As Double, _
    High As Double, _
    Low As Double, _
    ClosePrice As Double, _
    Volume As Long, _
    SplitAdjustedPrice As Double _
)
    pDateOfData = DateOfData
    pOpen = OpenPrice
    pHigh = High
    pLow = Low
    pClose = ClosePrice
    pVolume = Volume
    pSplitAdjustedPrice = SplitAdjustedPrice
End Sub
