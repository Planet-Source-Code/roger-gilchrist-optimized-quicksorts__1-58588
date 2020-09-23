Attribute VB_Name = "ModQuickSortDouble"
Option Explicit

Public Sub QuickSortDouble(Larray() As Double, _
                           L As Long, _
                           R As Long, _
                           bDir As Boolean)

  'Use on Double, Single, Long, Integer or Byte data
  'Do not use on Strings/Variant
  
  Dim I As Long
  Dim J As Long
  Dim X As Double

  I = L
  J = R
  X = Larray((L + R) / 2)
  Do While (I <= J)
    If bDir Then
      Do While (Larray(I) < X And I < R)
        I = I + 1
      Loop
      Do While (X < Larray(J) And J > L)
        J = J - 1
      Loop
     Else
      Do While (Larray(I) > X And I < R)
        I = I + 1
      Loop
      Do While (X > Larray(J) And J > L)
        J = J - 1
      Loop
    End If
    If I <= J Then
      SwapDouble Larray(I), Larray(J)
      I = I + 1
      J = J - 1
    End If
  Loop
  If L < J Then
    QuickSortDouble Larray, L, J, bDir
  End If
  If I < R Then
    QuickSortDouble Larray, I, R, bDir
  End If

End Sub

Private Sub SwapDouble(Dbl1 As Double, _
                       Dbl2 As Double)

  Dim Dbl3 As Double

  Dbl3 = Dbl1
  Dbl1 = Dbl2
  Dbl2 = Dbl3

End Sub

':)Code Fixer V2.9.1 (31/01/2005 9:01:03 PM) 1 + 56 = 57 Lines Thanks Ulli for inspiration and lots of code.

