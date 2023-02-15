VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   2400
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "stop"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Private Sub cmdstop_Click()
If flag = False Then
flag = True
cmdstop.Caption = "stop"
Else
flag = False
cmdstop.Caption = "start"
End If
End Sub

Private Sub Form_Load()
cmdstop.Caption = "stop"
flag = True
End Sub

Private Sub Image1_Click()
If flag = True Then
Image1.Top = Image1.Top + 100
Image1.Left = Image1.Left + 100
End If
End Sub
