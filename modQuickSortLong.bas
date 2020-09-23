Attribute VB_Name = "modQuickSortLong"
Option Explicit

Public Sub QuickSortLong(Larray() As Long, _
                         L As Long, _
                         R As Long, _
                         bDir As Boolean)

  'Use Long, Integer or Byte data
  'Do not use on Strings/Variant, Double, Single
  
  Dim I As Long
  Dim J As Long
  Dim X As Long

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
      SwapLong Larray(I), Larray(J)
      I = I + 1
      J = J - 1
    End If
  Loop
  If L < J Then
    QuickSortLong Larray, L, J, bDir
  End If
  If I < R Then
    QuickSortLong Larray, I, R, bDir
  End If

End Sub

Private Sub SwapLong(Lng1 As Long, _
                     Lng2 As Long)

  Dim Lng3 As Long

  Lng3 = Lng1
  Lng1 = Lng2
  Lng2 = Lng3

End Sub

':)Code Fixer V2.9.1 (31/01/2005 9:01:01 PM) 1 + 56 = 57 Lines Thanks Ulli for inspiration and lots of code.

