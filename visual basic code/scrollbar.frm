VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtinput 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblblue 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblgreen 
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblred 
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblcaption 
      Caption         =   "Change the color  of text"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
HScroll1.Min = 0
HScroll2.Min = 0
HScroll3.Min = 0
HScroll1.Max = 255
HScroll2.Max = 255
HScroll3.Max = 255
HScroll1.LargeChange = 50
HScroll2.LargeChange = 50
HScroll3.LargeChange = 50
End Sub

Private Sub HScroll1_Change()
txtinput.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
lblred = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
txtinput.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
lblgreen = HScroll2.Value
End Sub

Private Sub HScroll3_Change()
txtinput.ForeColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
lblblue = HScroll3.Value
End Sub

