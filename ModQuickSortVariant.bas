Attribute VB_Name = "ModQuickSortVariant"
Option Explicit

Public Sub QuickSortVariant(Larray As Variant, _
                            L As Long, _
                            R As Long, _
                            bDir As Boolean)

  'very slow
  'do not use on numeric data
  '
  
  Dim I As Long
  Dim J As Long
  Dim X As Variant

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
      SwapVariant Larray(I), Larray(J)
      I = I + 1
      J = J - 1
    End If
  Loop
  If L < J Then
    QuickSortVariant Larray, L, J, bDir
  End If
  If I < R Then
    QuickSortVariant Larray, I, R, bDir
  End If

End Sub

Private Sub SwapVariant(var1 As Variant, _
                        var2 As Variant)

  Dim var3 As Variant

  var3 = var1
  var1 = var2
  var2 = var3

End Sub

':)Code Fixer V2.9.1 (31/01/2005 9:01:06 PM) 1 + 57 = 58 Lines Thanks Ulli for inspiration and lots of code.

