Function f(a, b)
  If IsNumeric(a) And IsNumeric(b) Then
    c = a + b
  ElseIf IsNumeric(a) Then
    c = a & " + Non-numeric input"
  ElseIf IsNumeric(b) Then
    c = "Non-numeric input + " & b
  Else
    c = "Error: Non-numeric inputs"
  End If
  f = c
End Function

MsgBox f(10, "hello")
MsgBox f("hello", 10)
MsgBox f("hello", "world")
