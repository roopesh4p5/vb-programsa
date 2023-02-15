VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   ScaleHeight     =   5145
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstop 
      Caption         =   "stop"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   2280
      Picture         =   "move an image.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
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

Private Sub Timer1_Timer()
If flag = True Then
Image1.Top = Image1.Top + 100
Image1.Left = Image1.Left + 100
End If
End Sub
