Option Explicit

' length
Private pLength As Integer
' values array
Private pValues As Collection

Public Sub InitiateProperties( _
    Length As Integer _
    )
    pLength = Length
    Set pValues = New Collection
End Sub

' add data
Public Sub AddDataPoint(dataPoint As Double)
    If (pValues.Count >= pLength) Then pValues.Remove 1
    pValues.Add dataPoint
End Sub

' is initialized (has necessary data points
Public Function IsInitialized() As Boolean
    If pValues.Count = pLength Then IsInitialized = True
End Function

' value (average of all the data points
Public Function Value() As Double
    Dim sum As Double
    Dim v As Variant
    For Each v In pValues
        sum = sum + v
    Next v
    Value = sum / pValues.Count
End Function

