Function f(x)
  If x > 0 Then
    f = x^2
  ElseIf x = 0 Then
    f = 0
  Else
    f = -x^2
  End If
End Function

MsgBox f(2) ' Output: 4
MsgBox f(0) ' Output: 0
MsgBox f(-2) ' Output: -4