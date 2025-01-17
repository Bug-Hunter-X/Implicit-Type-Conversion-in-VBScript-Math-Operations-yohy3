Function f(a, b)
  If IsNumeric(a) And IsNumeric(b) Then
    c = a + b
  Else
    c = "Error: Non-numeric input"
  End If
  f = c
End Function

MsgBox f(10, "hello")