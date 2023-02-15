VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form3"
   ScaleHeight     =   5235
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmddecode 
      Caption         =   "decode"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdencode 
      Caption         =   "encode"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtinput 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbldecrypt 
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblencrypt 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "decrypted string"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "encripted string"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "enter a string"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "encriptio of a string"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SLEN As String
Dim KEY As String
Dim ORIGINAL As String
Dim ENSTRING As String
Dim DESTRING As String

Private Sub cmddecode_Click()
Dim KEYCH As String
DESTRING = ""
For I = 1 To SLEN
CH = Mid(ENSTRING, I, 1)
KEYCH = Mid(KEY, I, 1)
CH = Chr(Asc(CH) - Val(KEYCH))
DESTRING = DESTRING + CH
Next I
lbldecrypt.Caption = DESTRING

End Sub

Private Sub cmdencode_Click()
Dim X
Dim CH As String
KEY = ""
ENSTRING = ""
ORIGINAL = Trim(txtinput.Text)
SLEN = Len(ORIGINAL)
For I = 1 To SLEN
X = Int(Rnd() * 10)
CH = Mid(ORIGINAL, I, 1)
CH = Chr((Asc(CH) + X))
ENSTRING = ENSTRING + CH
KEY = KEY + LTrim(Str(X))
Next I
lblencrypt.Caption = ENSTRING
End Sub

Private Sub cmdexit_Click()
End
End Sub

Private Sub Form_Load()
Randomize
txtinput = ""
End Sub

