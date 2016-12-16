Function Max(ParamArray TheValues() As Variant) As Variant
'Return the "highest" of the values

Dim intLoop As Integer
Dim varCurrentMax As Variant
  varCurrentMax = TheValues(LBound(TheValues))
  For intLoop = LBound(TheValues) + 1 To UBound(TheValues)
    If TheValues(intLoop) > varCurrentMax Then
      varCurrentMax = TheValues(intLoop)
    End If
  Next intLoop
  
   Max = varCurrentMax

End Function

Function Min(ParamArray TheValues() As Variant) As Variant
'Return the "highest" of the values

Dim intLoop As Integer
Dim varCurrentMin As Variant
  varCurrentMax = TheValues(LBound(TheValues))
  For intLoop = LBound(TheValues) + 1 To UBound(TheValues)
    If TheValues(intLoop) < varCurrentMax Then
      varCurrentMin = TheValues(intLoop)
    End If
  Next intLoop
  
  Min = varCurrentMin

End Function
 
