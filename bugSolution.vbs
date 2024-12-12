Function f(a, b)
  'Convert inputs to numbers to avoid type mismatch
  Dim aNum, bNum
  aNum = CDbl(a)
  bNum = CDbl(b)
  If aNum > bNum Then
    MsgBox "a is greater than b"
  ElseIf aNum < bNum Then
    MsgBox "a is less than b"
  Else
    MsgBox "a is equal to b"
  End If
end function

' Calling the function with data type conversion
f("10", 20)