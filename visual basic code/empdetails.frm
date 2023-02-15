VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "employee  pay"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   360
         Top             =   3240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\Database13.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\Database13.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Table1"
         Caption         =   ""
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
         Caption         =   "exit"
         Height          =   255
         Left            =   5760
         TabIndex        =   20
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton cmdcalc 
         Caption         =   "calculate"
         Height          =   375
         Left            =   3360
         TabIndex        =   19
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtgross 
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtdeduct 
         Height          =   285
         Left            =   6120
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txthra 
         Height          =   285
         Left            =   5880
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtda 
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtbasic 
         DataField       =   "EMPBASIC"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtdesign 
         DataField       =   "EMPDESIGNATION"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtname 
         DataField       =   "EMPNAME"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtempid 
         DataField       =   "EMPID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lbldisplay 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "GROSSPAY"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "DEDUCTION"
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "HRA"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "DA"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "basic pay"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "designation"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "empname"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "empid"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "employee  pay calculation"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   0
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcalc_Click()
Dim basic As Double
Dim da As Double
Dim hra As Double
Dim deduct As Double
Dim gross As Double
Dim net As Double
basic = Val(txtbasic.Text)
Select Case basic
Case Is >= 50000
    da = 30 / 100 * basic
    hra = 15 / 100 * basic
    deduct = 12.5 / 100 * basic
    
Case Is >= 20000
    da = Str(25 / 100 * basic)
    hra = 12 / 100 * basic
    deduct = 8 / 100 * basic
Case Is >= 10000
    da = 15 / 100 * basic
    hra = 5 / 100 * basic
    deduct = 5 / 100 * basic
Case Else
    da = 10 / 100 * basic
    hra = 5 / 100 * basic
    deduct = 0
End Select

gross = basic + da + hra
net = gross - deduct
txtda.Text = da
txthra.Text = hra
txtdeduct.Text = deduct
txtgross = gross
lbldisplay.Caption = "NET SALARY=" + Str(net)
End Sub

Private Sub cmdexit_Click()
End
End Sub

