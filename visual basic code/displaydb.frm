VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "displaydb.frx":0000
      Height          =   1335
      Left            =   3480
      TabIndex        =   17
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "txtid"
         Caption         =   "txtid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "txtname"
         Caption         =   "txtname"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "txtrate"
         Caption         =   "txtrate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "txtdate"
         Caption         =   "txtdate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "txtunit"
         Caption         =   "txtunit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   120
      Top             =   4080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\displaydb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\displaydb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
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
      Caption         =   "exit"
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "update"
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdmodify 
      Caption         =   "modify"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtunit 
      DataField       =   "txtunit"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtdate 
      DataField       =   "txtdate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtrate 
      DataField       =   "txtrate"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      DataField       =   "txtname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtid 
      DataField       =   "txtid"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblmsg 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "UNIT OF MEASURE"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "MFD DATA"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "RATE PER UNIT"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "ITEM NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "ITEM ID"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "itemdetails"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Dim no As Integer
Adodc1.Refresh
lblmsg.Caption = "enter your details & press update"
lblmsg.Visible = True
If Adodc1.Recordset.BOF = True Then
no = 1
Else
Adodc1.Recordset.MoveLast
no = Adodc1.Recordset(0)
no = no + 1
End If
Adodc1.Recordset.AddNew
txtid.Text = no
txtdate.Text = no
cmdadd.Enabled = False
txtname.SetFocus
cmdupdate.Enabled = True
End Sub
Private Sub cmddelete_Click()
Dim ans
DataGrid1.Enabled = True
ans = MsgBox("are you sure? do you want to delete this record?", vbYesNo + vbQuestion, "confirmation")
If ans = vbYes Then
Adodc1.Recordset.Delete
MsgBox "record deleted"
End If
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdmodify_Click()
Adodc1.Recordset.Update
MsgBox "recordset updated"
lblmsg.Caption = "select the required record,modify your details and press update"
lblmsg.Visible = True
txtname.SetFocus
cmdupdate.Enabled = True
cmdmodify.Enabled = False
End Sub
Private Sub cmdupdate_Click()
txtdate.Text = CDate(txtdate.Text)
Adodc1.Recordset.Update
MsgBox "record added/updated successfully"
cmdadd.Enabled = True
cmdmodify.Enabled = True
cmdupdate.Enabled = False
lblmsg.Visible = False
End Sub
Private Sub Form_Load()
cmdupdate.Enabled = False
DataGrid1.Enabled = False
End Sub

