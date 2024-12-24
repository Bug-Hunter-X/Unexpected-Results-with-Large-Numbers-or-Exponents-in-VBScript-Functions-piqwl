Function f(x)
  Dim result
  If IsNumeric(x) Then
    If x > 0 Then
      result = x * x
    ElseIf x = 0 Then
      result = 0
    Else
      result = -1 * (x * x)
    End If 
  Else
    result = "Invalid Input: Not a Number"
  End If  
  f = result
End Function

MsgBox f(2) ' Output: 4
MsgBox f(0) ' Output: 0
MsgBox f(-2) ' Output: -4
MsgBox f("a") 'Output: Invalid Input: Not a Number