VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmddispense 
      Caption         =   "Dispense snack"
      Height          =   975
      Left            =   5760
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtbill 
      Height          =   735
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtchoice 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Enter your choice"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblmsg 
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   240
      Picture         =   "vending.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "4"
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "3"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "2-kur"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "1-lays"
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   360
      Picture         =   "vending.frx":91560
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   360
      Picture         =   "vending.frx":128EC4
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "vending.frx":1FF9F2
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "VENDING MACHINE"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim snack
Dim rate
Dim sel_item
Dim sel_rate
Private Sub cmddispense_Click()
If txtchoice.Text = "" Then
MsgBox "enter your choice"
txtchoice.SetFocus
Else
sel_item = snack(Val(txtchoice.Text) - 1)
sel_rate = rate(Val(txtchoice.Text) - 1)
lblmsg.Caption = sel_item
lblmsg.Visible = True
txtbill.Text = txtbill.Text + sel_item + "" + Str(sel_rate)
txtbill.Visible = True
End If
End Sub
Private Sub Command1_Click()
End
End Sub
Private Sub Form_activate()
txtchoice.SetFocus
txtbill.Enabled = False
End Sub
Private Sub Form_Load()
lblmsg.Visible = False
snack = Array("lays", "kurkure", "dairymilk", "kitkat")
rate = Array(10, 10, 5, 15)
txtbill.Visible = False
End Sub

Private Sub Image1_Click()
txtchoice.Text = Index + 1
End Sub
