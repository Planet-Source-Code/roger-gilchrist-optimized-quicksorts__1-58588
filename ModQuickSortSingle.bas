Attribute VB_Name = "ModQuickSortSingle"
Option Explicit

Public Sub QuickSortSingle(Larray() As Single, _
                           L As Long, _
                           R As Long, _
                           bDir As Boolean)

  'Use on Single , Long,Integer or Byte
  'Do not use on Strings/Variant or Double
  
  Dim I As Long
  Dim J As Long
  Dim X As Single

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
      SwapSingle Larray(I), Larray(J)
      I = I + 1
      J = J - 1
    End If
  Loop
  If L < J Then
    QuickSortSingle Larray, L, J, bDir
  End If
  If I < R Then
    QuickSortSingle Larray, I, R, bDir
  End If

End Sub

Private Sub SwapSingle(Sng1 As Single, _
                       Sng2 As Single)

  Dim Sng3 As Single

  Sng3 = Sng1
  Sng1 = Sng2
  Sng2 = Sng3

End Sub

':)Code Fixer V2.9.1 (31/01/2005 9:01:04 PM) 1 + 56 = 57 Lines Thanks Ulli for inspiration and lots of code.

