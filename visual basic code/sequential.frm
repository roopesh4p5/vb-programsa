VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   17
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "edit"
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   3735
      Begin VB.CommandButton cmdclear 
         Caption         =   "clear"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdread 
         Caption         =   "read"
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdwrite 
         Caption         =   "write"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PERMISSION"
      Height          =   3495
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   2295
      Begin VB.CommandButton cmdexit 
         Caption         =   "exit"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdopenread 
         Caption         =   "open for read"
         Height          =   735
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdopenwrite 
         Caption         =   "OPEN FOR WRITE"
         Height          =   855
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label5 
      Caption         =   "PHONE"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "PINCODE"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "CITY"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ADDRESS"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "NAME"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnt As Integer
Dim fileno As Integer


Private Sub cmdclear_Click()
Call clearall
End Sub

Private Sub cmdexit_Click()
Close
Unload Me
End Sub

Private Sub cmdopenread_Click()
Close
fileno = FreeFile
Open "address.txt" For Input As fileno
cmdopenwrite.Enabled = True
cmdopenread.Enabled = False
cmdread.Enabled = True
cmdwrite.Enabled = False
cmdread.SetFocus
End Sub

Private Sub cmdopenwrite_Click()
Close fileno = FreeFile
Open "address.txt" For Append As fileno
Text1(0).SetFocus
cmdopenwrite.Enabled = False
cmdopenread.Enabled = True
cmdwrite.Enabled = True
cmdread.Enabled = False
End Sub

Private Sub cmdread_Click()
On Error GoTo errortrap
Dim fieldcontent
If Not EOF(fileno) Then
Call clearall
For cnt = 0 To 4
Input #fileno, fieldcontent
Text1(cnt) = fieldcontent
Next cnt
End If
If EOF(fileno) Then
MsgBox "end of file"
cmdexit.SetFocus
End If
Exit Sub
Error trap:
MsgBox Err.Description
End Sub

Private Sub cmdwrite_Click()
For cnt = 0 To 4
If Text1(cnt) = "" Then
MsgBox "enter valid data"
Text1(cnt).SetFocus
Exit Sub
End If
Write #fileno, Text1(cnt);
Next cnt
Call clearall
End Sub

Private Sub clearall()
For cnt = 0 To 4
Text1(cnt) = ""
Next cnt
Text1(0).SetFocus
End Sub
