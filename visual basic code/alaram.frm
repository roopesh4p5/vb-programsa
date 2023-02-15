VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3720
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   2400
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "stop"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "setalarm"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtalarm 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      Picture         =   "alaram.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lbltime 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Alarm time"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "current time"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim min1 As Integer
Dim min2 As Integer

Private Sub cmdok_Click()
Timer1.Enabled = True
cmdok.Enabled = False
cmdset.Enabled = True
Dim alarm As Integer
min2 = Minute(txtalarm.Text)
End Sub

Private Sub cmdset_Click()
txtalarm.Enabled = True
txtalarm.SetFocus
cmdok.Enabled = True
cmdset.Enabled = False
min2 = 0
End Sub

Private Sub cmdstop_Click()
Timer2.Enabled = False
Image1.Visible = False
Timer1.Enabled = False
End Sub

Private Sub Form_Load()
Timer2.Enabled = False
lbltime.Caption = Time
Timer1.Interval = 1000
lbltime.AutoSize = True
txtalarm.Enabled = False
cmdok.Enabled = False
Image1.Visible = False
flag = False
min2 = 0
End Sub



Private Sub Timer1_Timer()
lbltime.Caption = Time
min1 = Minute(Now)
If min1 = min2 Then
Timer2.Enabled = True
Timer2.Enabled = 1000
End If
End Sub

Private Sub Timer2_Timer()
If min1 = min2 And Image1.Visible = False Then
Image1.Visible = True
Else
Image1.Visible = False
End If
End Sub
