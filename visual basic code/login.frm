VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   1680
      Top             =   3360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1720
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\logindb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\logindb.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton cmdcancel 
      Caption         =   "cancel"
      Height          =   735
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "login"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtpass 
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtuser 
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New adodb.Connection
Dim RS As New adodb.Recordset

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdok_Click()
Set conn = New Connection
Set RS = New Recordset
conn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\logindb.mdb;Persist Security Info=False"
RS.Open " select * from login where username='" & txtuser.Text & "' and password='" & txtpass.Text & "'", conn

If RS.EOF = False Then
txtuser.Text = ""
txtpass.Text = ""
MsgBox "login sucessfull"
RS.Close
Else
MsgBox "incorrect username/password login denied"
txtuser.Text = ""
txtpass.Text = ""
txtuser.SetFocus
RS.Close
End If
End Sub

Private Sub Form_Load()
txtpass.PasswordChar = "*"
txtpass.Text = ""
End Sub
