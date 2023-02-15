VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OLE OLE3 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OLE OLE2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub text1_()

End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.OLEDrag

End Sub




Private Sub Text1_OLECompleteDrag(Effect As Long)
MsgBox "Returned OLE effect:" & Effect
End Sub

Private Sub text1_OLEDStartDrag(Data As DataObject, allowedeffects As Long)
Data.SetData Text1.Text, vbCFText

End Sub

Private Sub Text2_()

End Sub

Private Sub Text2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.Text = Data.GetData(vbCFText)
End Sub
