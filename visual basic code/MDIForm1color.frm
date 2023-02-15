VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunew 
         Caption         =   "&new"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "&open"
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "&Format"
      Begin VB.Menu mnuforecolor 
         Caption         =   "&forecolor"
      End
      Begin VB.Menu mnubackcolor 
         Caption         =   "&backcolor"
      End
   End
   Begin VB.Menu mnufont 
      Caption         =   "&Font"
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnubackcolor_Click()
CommonDialog1.ShowColor
Form1.Text1.BackColor = CommonDialog1.Color
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufont_Click()
'CommonDialog1.Flags = cdlCFBoth Or cdlFEffects
CommonDialog1.ShowFont
Form1.Text1.FontName = CommonDialog1.FontName
Form1.Text1.FontBold = CommonDialog1.FontBold
Form1.Text1.FontItalic = CommonDialog1.FontItalic
Form1.Text1.FontSize = CommonDialog1.FontSize
Form1.Text1.FontUnderline = CommonDialog1.FontUnderline
Form1.Text1.FontStrikethru = CommonDialog1.FontStrikethru
End Sub

Private Sub mnuforecolor_Click()
CommonDialog1.ShowColor
Form1.Text1.ForeColor = CommonDialog1.Color
End Sub

Private Sub mnunew_Click()
Load Form1
Form1.Show
End Sub

Private Sub mnuopen_Click()
CommonDialog1.ShowOpen
End Sub
