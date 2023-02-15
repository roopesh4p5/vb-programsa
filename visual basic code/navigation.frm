VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form form1 
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "class"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "rollno"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   120
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "navigation.frx":0000
      Height          =   1095
      Left            =   6360
      TabIndex        =   15
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1931
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
      ColumnCount     =   4
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
         DataField       =   "rollno"
         Caption         =   "rollno"
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
         DataField       =   "name"
         Caption         =   "name"
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
         DataField       =   "class"
         Caption         =   "class"
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
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   720
      Top             =   4320
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\navigation.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Projector\Documents\navigation.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "refresh"
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "exit"
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "save"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "update"
      Height          =   495
      Left            =   8640
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "delete"
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "add"
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "previous"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "next"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "last"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "first"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "manipulation method"
      Height          =   1575
      Left            =   6000
      TabIndex        =   4
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Navigation method"
      Height          =   1575
      Left            =   840
      TabIndex        =   3
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "class"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "name"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Rollno"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
Dim rno As Integer
'On Error GoTo errmsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
rno = Adodc1.Recordset("rollno") + 1
Adodc1.Recordset.AddNew
Text1.Text = 1
Adodc1.Recordset.AddNew
End Sub

Private Sub cmddelete_Click()
On Error GoTo errmsg
Dim wish As Integer
wish = MsgBox("are you sure to delete the record(Y/N)?", vbYesNo + vbQuestion)
If wish = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
Adodc1.Recordset.MovePrevious
End If
MsgBox "record deleted"
End If
Exit Sub
errmsg:
MsgBox "record not found"
End Sub
Private Sub cmdexit_Click()
End
End Sub
Private Sub cmdfirst_Click()
Adodc1.Recordset.MoveFirst
End Sub
Private Sub cmdlast_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveLast
MsgBox "this is a last record in the database"
End If
End Sub

Private Sub cmdprevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
MsgBox "this is the first record in the database"
End If
End Sub

Private Sub cmdsave_Click()
On Error GoTo errmsg
Adodc1.Recordset.Update
MsgBox "record saved"
errmsg:
MsgBox "duplicate record number"
Adodc1.Recordset.CancelUpdate
End Sub

Private Sub cmdupdate_Click()
Adodc1.Recordset.Update
MsgBox "record altered"
End Sub

