VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10470
   LinkTopic       =   "Form3"
   ScaleHeight     =   7005
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   1215
      Left            =   960
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "bill"
      Height          =   2655
      Left            =   4440
      TabIndex        =   9
      Top             =   3960
      Width           =   3855
      Begin VB.Label lblamt 
         Height          =   735
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblamt 
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblname 
         Height          =   735
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Sweet Fresh Limited Bangalore"
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdbill 
      Caption         =   "print bill"
      Height          =   855
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmddispense 
      Caption         =   "Dispense"
      Height          =   615
      Left            =   4560
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtno 
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "snacks"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      Begin VB.Image Image4 
         Height          =   375
         Left            =   1200
         Top             =   2280
         Width           =   735
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   960
         Top             =   1440
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   960
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   1080
         Picture         =   "vendingtxtbook.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
      Begin VB.Label label4 
         Caption         =   "4.rosagulla"
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "3.obattu"
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label label2 
         Caption         =   "2.burfi"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label label1 
         Caption         =   "1.dryfruit"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label lblsnack 
      Height          =   495
      Left            =   6360
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Enter snack NO:"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdbill_Click()
cmdbill.Enabled = False
If txtno = "" Then
MsgBox "please enter the snack num"
txtno.SetFocus
Exit Sub
End If


Frame2.Visible = True
If txtno = 1 Then
lblname.Caption = "dry fruits"
lblamt.c = 160
ElseIf txtno = 2 Then
lblname.Caption = "burfi"
lblamt.Caption = "200"
ElseIf txtno = 3 Then
lblname.Caption = "obattu"
lblamt.Caption = 60
ElseIf txtno = 4 Then
lblname.Caption = "rosagulla"
lblamt.Caption = 250
Else
MsgBox "snack does not exit!!!"
txtno = ""
txtno.SetFocus
End If

End Sub

Private Sub cmddispense_Click()
Frame2.Visible = False
cmdbill.Enabled = True
If txtno = "" Then
MsgBox "Please enter the snack num"
txtno.SetFocus
Exit Sub
End If


If txtno = 1 Then
lblsnack.Caption = "dry fruits"
ElseIf txtno = 2 Then
lblsnack.Caption = "burfi"
ElseIf txtno = 3 Then
lblsnack.Caption = "obattu"
ElseIf txtno = 4 Then
lblsnack.Caption = "rosagulla"
Else: MsgBox "snack does not exit "
lblsnack.Caption = ""
txtno = ""
txtno.SetFocus
End If
End Sub

Private Sub Command2_Click()
End
End Sub
