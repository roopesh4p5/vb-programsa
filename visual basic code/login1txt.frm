VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   3120
      Top             =   4200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Desktop\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Desktop\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1815
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      Begin VB.TextBox txtpassword 
         DataField       =   "password"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtusername 
         DataField       =   "username"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         DataField       =   " "
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s1 As String
Dim s2 As String

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdlogin_Click()
If txtusername = "" Or txtpassword = "" Then
MsgBox "Please enter all details"
txtusername.SetFocus
Exit Sub
End If
s1 = txtusername.Text
s2 = txtpassword.Text
Adodc1.RecordSource = "select * from login where username='" + s1 + "'and password='" + s2 + "'"
Adodc1.Refresh
If adodoc1.Recordset.recordcount = 0 Then
MsgBox "Invalid login:"
Else
MsgBox "valid user"
Exit Sub
End If
txtusername.Text = ""
txtpassword.Text = ""
txtusername.SetFocus
End Sub

Private Sub txtusername_Change()

End Sub
