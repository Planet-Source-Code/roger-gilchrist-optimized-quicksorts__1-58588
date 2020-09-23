VERSION 5.00
Begin VB.Form frmQuickSort 
   Caption         =   "QuickSorts Demo"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "Variant >"
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "String >"
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "Double >"
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Big Double"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   1230
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Big Variant"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Big Long"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   1980
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Big Single"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   1605
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "big String"
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   615
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "SIngle >"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "Generate"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "Long >"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox lstQuickSort 
      Height          =   6495
      Index           =   2
      Left            =   5670
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstQuickSort 
      Height          =   6495
      Index           =   1
      Left            =   4455
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lstQuickSort 
      Height          =   6495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   7
      X1              =   6840
      X2              =   6840
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   4455
      Left            =   7080
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   3855
      Left            =   1200
      TabIndex        =   7
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Random                                                                                    Ascending          Descending"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmQuickSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ASize     As Long = 32

Private Sub cmdQuickSort_Click(Index As Integer)

  Dim I             As Long
  Dim arrV(ASize)   As Variant
  Dim arrStr(ASize) As String
  Dim arrD(ASize)   As Double
  Dim arrS(ASize)   As Single
  Dim arrL(ASize)   As Long

  lstQuickSort(1).Clear
  lstQuickSort(2).Clear
  Select Case Index
   Case 0
    ReSet
   Case 1
    For I = 0 To ASize
      arrL(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortLong arrL, 0, UBound(arrL), True
    For I = 0 To ASize
      lstQuickSort(1).AddItem arrL(I)
    Next I
    For I = 0 To ASize
      arrL(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortLong arrL, 0, UBound(arrL), False
    For I = 0 To ASize
      lstQuickSort(2).AddItem arrL(I)
    Next I
   Case 2
    For I = 0 To ASize
      arrD(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortDouble arrD, 0, UBound(arrD), True
    For I = 0 To ASize
      lstQuickSort(1).AddItem arrD(I)
    Next I
    For I = 0 To ASize
      arrD(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortDouble arrD, 0, UBound(arrD), False
    For I = 0 To ASize
      lstQuickSort(2).AddItem arrD(I)
    Next I
   Case 3
    For I = 0 To ASize
      arrS(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortSingle arrS, 0, UBound(arrS), True
    For I = 0 To ASize
      lstQuickSort(1).AddItem arrS(I)
    Next I
    For I = 0 To ASize
      arrS(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortSingle arrS, 0, UBound(arrS), False
    For I = 0 To ASize
      lstQuickSort(2).AddItem arrS(I)
    Next I
   Case 4
    For I = 0 To ASize
      arrStr(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortString arrStr, 0, UBound(arrStr), True
    For I = 0 To ASize
      lstQuickSort(1).AddItem arrStr(I)
    Next I
    For I = 0 To ASize
      arrStr(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortString arrStr, 0, UBound(arrStr), False
    For I = 0 To ASize
      lstQuickSort(2).AddItem arrStr(I)
    Next I
   Case 5
    For I = 0 To ASize
      arrV(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortVariant arrV, 0, UBound(arrV), True
    For I = 0 To ASize
      lstQuickSort(1).AddItem arrV(I)
    Next I
    For I = 0 To ASize
      arrV(I) = lstQuickSort(0).List(I)
    Next I
    QuickSortVariant arrV, 0, UBound(arrV), False
    For I = 0 To ASize
      lstQuickSort(2).AddItem arrV(I)
    Next I
  End Select

End Sub

Private Sub Command1_Click()

  Dim arrBig(1000000) As String
  Dim T               As Double

  Dim I               As Long
  Command1.Enabled = False
  Command1.Caption = "Generating"
  For I = 0 To 1000000
    arrBig(I) = rndWord
  Next I
  Command1.Caption = "Sorting"
  T = Timer
  QuickSortString arrBig, 0, UBound(arrBig), False
  Command1.Caption = "Big String " & Timer - T
  Command1.Enabled = True

End Sub

Private Sub Command2_Click()

  Dim arrBig(1000000) As Single
  Dim T               As Double

  Dim I               As Long
  Command2.Enabled = False
  Command2.Caption = "Generating"
  For I = 0 To 1000000
    arrBig(I) = Rnd * 1000000
  Next I
  Command2.Caption = "Sorting"
  T = Timer
  QuickSortSingle arrBig, 0, UBound(arrBig), False
  Command2.Caption = "Big Single " & Timer - T
  Command2.Enabled = True

End Sub

Private Sub Command3_Click()

  Dim arrBig(1000000) As Long
  Dim T               As Double

  Dim I               As Long
  Command3.Enabled = False
  Command3.Caption = "Generating"
  For I = 0 To 1000000
    arrBig(I) = CLng(Rnd * 1000000)
  Next I
  Command3.Caption = "Sorting"
  T = Timer
  QuickSortLong arrBig, 0, UBound(arrBig), False
  Command3.Caption = "Big Long " & Timer - T
  Command3.Enabled = True

End Sub

Private Sub Command4_Click()

  Dim arrBig(1000000) As Variant
  Dim T               As Double

  Dim I               As Long
  Command4.Enabled = False
  Command4.Caption = "Generating"
  For I = 0 To 1000000
    arrBig(I) = rndWord
  Next I
  Command4.Caption = "Sorting"
  T = Timer
  QuickSortVariant arrBig, 0, UBound(arrBig), False
  Command4.Caption = "Big Variant " & Timer - T
  Command4.Enabled = True

End Sub

Private Sub Command5_Click()

  Dim arrBig(1000000) As Double
  Dim T               As Double

  Dim I               As Long
  Command5.Enabled = False
  Command5.Caption = "Generating"
  For I = 0 To 1000000
    arrBig(I) = Rnd * 1000000
  Next I
  Command5.Caption = "Sorting"
  T = Timer
  QuickSortDouble arrBig, 0, UBound(arrBig), False
  Command5.Caption = "Big Double " & Timer - T
  Command5.Enabled = True

End Sub

Private Sub Form_Load()

  Label2 = "QUICKSORT DEMO" & vbNewLine & _
           "These buttons are designed to show the effect of different varieties of QuickSort by using small arrays of data" & vbNewLine & _
           "The arrays are much too small for significant timing results." & vbNewLine & _
           "Numeric cannot be applied to Strings at all so are blocked in code." & vbNewLine & _
           "" & vbNewLine & _
           "Things to watch out for:" & vbNewLine & _
           "1: String/Variant QuickSorts on numeric data." & vbNewLine & _
           "2: Long QuickSort on Single/Double data."
  Label3 = "QuickSort is very fast especially if you use optimised routines." & vbNewLine & _
           "The buttons above creates million member arrays, quicksort them and return the time taken." & vbNewLine & _
           "On my machine the speeds (in seconds) are" & vbNewLine & _
           "Sort        IDE      Compiled" & vbNewLine & _
           "Variant     55         40" & vbNewLine & _
           "String       25         12" & vbNewLine & _
           "Numeric   12          2" & vbNewLine & _
           "These are approx because each test generates its own random array which may be more or less suitable to sort." & vbNewLine & _
           "" & vbNewLine & _
           "Note: QuickSort may perform badly on data that is nearly sorted. It is advised that you don't apply it twice to the same data, but you could use the ascending/ descending sorts to do avoid this condition"
  ReSet

End Sub

Private Sub ReSet()

  Dim I As Long

  cmdQuickSort(1).Enabled = True
  cmdQuickSort(2).Enabled = True
  cmdQuickSort(3).Enabled = True
  cmdQuickSort(4).Enabled = True
  cmdQuickSort(5).Enabled = True
  lstQuickSort(0).Clear
  Select Case Rnd
   Case Is < 0.33
    For I = 0 To ASize
      lstQuickSort(0).AddItem Rnd * 100
    Next I
   Case Is < 0.66
    For I = 0 To ASize
      lstQuickSort(0).AddItem Int(Rnd * 100)
    Next I
   Case Else
    For I = 0 To ASize
      lstQuickSort(0).AddItem rndWord
    Next I
    cmdQuickSort(1).Enabled = False
    cmdQuickSort(2).Enabled = False
    cmdQuickSort(3).Enabled = False
  End Select

End Sub

Private Function rndWord() As String

  Dim I  As Long
  Dim ca As Long

  ca = IIf(Rnd > 0.5, 97, 65)
  For I = 1 To Int(Rnd * 4) + 3
    rndWord = rndWord & Chr$(Int(Rnd * 26) + ca)
  Next I

End Function

':)Code Fixer V2.9.1 (31/01/2005 9:01:00 PM) 2 + 246 = 248 Lines Thanks Ulli for inspiration and lots of code.

