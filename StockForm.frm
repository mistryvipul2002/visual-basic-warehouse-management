VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form StockForm 
   Caption         =   "Stock"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "PRODUCTCODE"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Exit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Search 
      Height          =   285
      Left            =   7320
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Edit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "StockForm.frx":0000
      Height          =   4095
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1080
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\WareHouseDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\WareHouseDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Stock"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5760
      Top             =   5400
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\WareHouseDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\WareHouseDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"StockForm.frx":0015
      Caption         =   "Adodc2"
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
   Begin VB.Label LabelSearch 
      Caption         =   "Search by Product Code:"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "StockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call initializeConnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainForm.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Delete_Click()
Adodc1.recordSet.Delete
End Sub

Private Sub Add_Click()
Adodc2.Refresh
If Adodc2.recordSet.EOF Then
    MsgBox "You cannot Add. Either Edit or Delete"
Else
    Me.Hide
    StockInfo.Show
    StockInfo.Mode = "Add"
    
    StockInfo.Text1.Visible = False
    StockInfo.Combo1.Visible = True
    
    Adodc2.recordSet.MoveFirst
    StockInfo.Combo1.Clear
    While Not Adodc2.recordSet.EOF
        StockInfo.Combo1.AddItem Adodc2.recordSet.Fields("PRODUCTCODE")
        Adodc2.recordSet.MoveNext
    Wend
    StockInfo.Combo1.ListIndex = 0
End If
End Sub

Private Sub Edit_Click()
Me.Hide
StockInfo.Show
StockInfo.Mode = "Edit"

StockInfo.Combo1.Visible = False
StockInfo.Text1.Visible = True

StockInfo.Text1.Text = Adodc1.recordSet.Fields("PRODUCTCODE")
StockInfo.str1 = StockInfo.Text1.Text

StockInfo.Text2.Text = Adodc1.recordSet.Fields("Quantity")
StockInfo.str2 = StockInfo.Text2.Text

Adodc1.recordSet.Delete
End Sub

Private Sub Search_Change()
If Search.Text <> "" Then
    If Not search2(Adodc1.recordSet, "PRODUCTCODE", Search.Text & "*") Then
        MsgBox "Not Found"
        Search.Text = ""
        Adodc1.recordSet.MoveFirst
    End If
End If
End Sub

