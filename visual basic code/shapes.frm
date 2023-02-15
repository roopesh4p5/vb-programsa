VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdline 
      Caption         =   "changeline"
      Height          =   615
      Left            =   480
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "changeshape"
      Height          =   615
      Left            =   4440
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtlinestyle 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtlinewidth 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtborder 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtstyle 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtshape 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "borderstyle(0-6)"
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   1800
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Shape           =   2  'Oval
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   240
      Shape           =   1  'Square
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "borderstyle(1-6)"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "borderwidth(0-5)"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "fillstyle(0-7)"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "shape(1-5)"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "shapecontrol"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdchange_Click()
If Val(txtshape) > 5 Then
MsgBox "invalid data"
txtshape = ""
txtshape.SetFocus
Exit Sub
Else
Shape1.Shape = Val(txtshape)
End If
If Val(txtstyle) > 7 Then
MsgBox "invalid data"
txtstyle = ""
txtstyle.SetFocus
Exit Sub
Shape1.FillStyle = Val(txtstyle)
End If
If Val(txtstyle) > 6 Then
MsgBox "invalid data"
txtborder = ""
txtborder.SetFocus
Exit Sub
Else
Shape1.BackStyle = Val(txtborder)
End If

End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdline_Click()
If Val(txtlinestyle) > 6 Then
MsgBox "invalid data"
txtlinestyle.SetFocus
Exit Sub
Else
Line1.BorderStyle = Val(txtlinestyle)
End If

End Sub

