VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "+"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   615
      Left            =   4560
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   44923
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   44916
   End
   Begin VB.CommandButton cmdage 
      Caption         =   "click here to know age"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblage 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "age"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "select DOB"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "age calculation usig DTpicker control"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdage_Click()
Dim bday As Integer
Dim bmonth As Integer
Dim byear As Integer
bday = DTPicker1.Day
bmonth = DTPicker1.Month
byear = DTPicker1.Year
If Year(Now) <= DTPicker1.Year Then
lblage.Caption = 0
ElseIf (bmonth < Month(Now) Or (bday < Day(Now))) Then
lblage.Caption = Abs((Year(Now) - DTPicker1.Year) - 1)
Else
lblage.Caption = Abs((Year(Now) - DTPicker1.Year))
End If

End Sub
