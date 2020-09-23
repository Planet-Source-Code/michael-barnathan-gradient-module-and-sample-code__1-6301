VERSION 4.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Gradients"
   ClientHeight    =   5310
   ClientLeft      =   1140
   ClientTop       =   1695
   ClientWidth     =   6690
   Height          =   5715
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6690
   Top             =   1350
   Width           =   6810
   Begin VB.CommandButton Command3 
      Caption         =   "Optimum Quality value"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3480
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   20
      Top             =   3120
      Value           =   100
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create gradient on form"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   6495
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   120
      ScaleHeight     =   435
      ScaleWidth      =   6435
      TabIndex        =   18
      Top             =   4800
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create gradient on picture"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   6495
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   10
      Top             =   1800
      Value           =   255
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   9
      Top             =   2160
      Value           =   255
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Index           =   2
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   8
      Top             =   2520
      Value           =   255
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   2
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Style:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quality: 100%"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "From:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Red: 255"
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   13
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Green: 255"
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   12
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Blue: 255"
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Blue: 0"
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Green: 0"
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red: 0"
      Height          =   195
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call CreateGradient(Picture3, HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value, HScroll2(0).Value, HScroll2(1).Value, HScroll2(2).Value, Combo1.ListIndex + 1, 0, (Picture3.Width / 100) * (HScroll3.Value), 0)
End Sub

Private Sub Command2_Click()
Call CreateGradient(Form1, HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value, HScroll2(0).Value, HScroll2(1).Value, HScroll2(2).Value, Combo1.ListIndex + 1, 0, (Form1.Width / 100) * (HScroll3.Value), 0)
End Sub


Private Sub Command3_Click()
If Combo1.ListIndex = 2 Then HScroll3.Value = 10: Else: HScroll3.Value = 5
End Sub


Private Sub Form_Load()
Combo1.AddItem "Horizontal"
Combo1.AddItem "Vertical"
Combo1.AddItem "Circular"
Combo1.ListIndex = 0
End Sub


Private Sub HScroll1_Change(Index As Integer)
Picture1.BackColor = RGB(HScroll1(0).Value, HScroll1(1).Value, HScroll1(2).Value)
Label1(0).Caption = "Red: " & SeperateColors(Picture1).Red
Label1(1).Caption = "Green: " & SeperateColors(Picture1).Green
Label1(2).Caption = "Blue: " & SeperateColors(Picture1).Blue
End Sub


Private Sub HScroll2_Change(Index As Integer)
Picture2.BackColor = RGB(HScroll2(0).Value, HScroll2(1).Value, HScroll2(2).Value)
Label2(0).Caption = "Red: " & SeperateColors(Picture2).Red
Label2(1).Caption = "Green: " & SeperateColors(Picture2).Green
Label2(2).Caption = "Blue: " & SeperateColors(Picture2).Blue
End Sub


Private Sub HScroll3_Change()
Label4.Caption = "Quality: " & HScroll3.Value & "%"
End Sub


Private Sub Label6_Click()

End Sub


