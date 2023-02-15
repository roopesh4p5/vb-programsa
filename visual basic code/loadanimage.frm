VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2520
      Picture         =   "loadanimage.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
End
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
Dim fname As String
fname = File1.Path + "\" + File1.FileName
Image1.Picture = LoadPicture(fname)
End Sub
