Attribute VB_Name = "ModQuickSortString"
Option Explicit

Public Sub QuickSortString(Larray() As String, _
                           L As Long, _
                           R As Long, _
                           bDir As Boolean)

  'faster than QuickSortVariant
  'do not use on numeric data
  
  Dim I As Long
  Dim J As Long
  Dim X As String

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
      SwapString Larray(I), Larray(J)
      I = I + 1
      J = J - 1
    End If
  Loop
  If L < J Then
    QuickSortString Larray, L, J, bDir
  End If
  If I < R Then
    QuickSortString Larray, I, R, bDir
  End If

End Sub

Private Sub SwapString(str1 As String, _
                       str2 As String)

  Dim str3 As String

  str3 = str1
  str1 = str2
  str2 = str3

End Sub

':)Code Fixer V2.9.1 (31/01/2005 9:01:05 PM) 1 + 56 = 57 Lines Thanks Ulli for inspiration and lots of code.

