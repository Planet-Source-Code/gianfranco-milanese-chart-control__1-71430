VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin chart.CtrChart CtrChart2 
      Height          =   2280
      Left            =   4650
      TabIndex        =   3
      Top             =   525
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4022
      BackColor       =   11200511
      MainBarColor    =   14320712
      ValueCount      =   10
      Highlight       =   255
      BgLine          =   -1  'True
      Space           =   8
   End
   Begin chart.CtrChart CtrChart3 
      Height          =   2220
      Left            =   -165
      TabIndex        =   2
      Top             =   3195
      Width           =   7470
      _ExtentX        =   5847
      _ExtentY        =   3678
      BackColor       =   43741
      MainBarColor    =   3158064
      ValueCount      =   30
      Highlight       =   14483677
      BgLine          =   0   'False
      Space           =   0
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   330
      Left            =   330
      Max             =   48
      Min             =   1
      TabIndex        =   1
      Top             =   2580
      Value           =   1
      Width           =   3375
   End
   Begin chart.CtrChart CtrChart1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   4180
      BackColor       =   -2147483629
      MainBarColor    =   -2147483645
      ValueCount      =   10
      Highlight       =   65535
      BgLine          =   -1  'True
      Space           =   8
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mValue(60)     As Double
Private mBarLbl(60)      As String
Private BarColor(60)     As Long

Private Sub Form_Load()

   Dim a   As Integer
   Dim max As Double

   max = 0
   For a = 1 To 60
      mValue(a) = Int(Rnd(1) * 10000) / 100
      mBarLbl(a) = Format$("01/" & CStr((a Mod 12) + 1) & "/01", "Mmm")
      If max < mValue(a) Then
         max = mValue(a)
      End If
      BarColor(a) = -1
   Next a
   CtrChart1.max = max
   BarColor(10) = &HFF00&
   HScroll1_Change

   For a = 1 To CtrChart2.ValueCount
      CtrChart2.Value(a) = Int(Rnd(1) * 100)
      CtrChart2.BarColor(a) = Int(Rnd(1) * 255) + &H100 * Int(Rnd(1) * 255) + &H10000 * Int(Rnd(1) * 255)
   Next a
   For a = 1 To CtrChart3.ValueCount
      CtrChart3.Value(a) = Int(Rnd(1) * 100)
   Next a
   
End Sub

Private Sub Form_Resize()

   CtrChart1.Move 0, 0, Me.ScaleWidth / 2, (Me.ScaleHeight / 2) - 330
   HScroll1.Move 0, CtrChart1.height, CtrChart1.width, 330
   CtrChart2.Move CtrChart1.width, 0, Me.ScaleWidth / 2, (Me.ScaleHeight / 2)
   CtrChart3.Move 0, CtrChart2.height, Me.ScaleWidth, (Me.ScaleHeight / 2)
   

End Sub

Private Sub HScroll1_Change()

   Dim a As Integer

   CtrChart1.Redraw = False
   CtrChart1.ValueCount = 12
   For a = HScroll1 To HScroll1 + 11
      With CtrChart1
         .Value(a + 1 - HScroll1) = mValue(a)
         .Text(a + 1 - HScroll1) = mBarLbl(a)
         .BarColor(a + 1 - HScroll1) = BarColor(a)
      End With 'CTRCHART1
   Next a
   CtrChart1.Redraw = True

End Sub

Private Sub HScroll1_Scroll()

   HScroll1_Change

End Sub

':)Code Fixer V3.0.9 (21/11/2008 17.28.02) 4 + 53 = 57 Lines Thanks Ulli for inspiration and lots of code.
