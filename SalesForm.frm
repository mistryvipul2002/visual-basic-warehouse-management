VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SalesForm 
   Caption         =   "Sales Form"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "CustomerID"
      DataSource      =   "Adodc4"
      Height          =   285
      Left            =   10560
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text2 
      DataField       =   "SRID"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   5160
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton View 
      Caption         =   "View"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "SRID"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Exit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Search 
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SalesForm.frx":0000
      Height          =   4095
      Left            =   600
      TabIndex        =   3
      Top             =   1080
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
            LCID            =   16393
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
            LCID            =   16393
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
      Left            =   1200
      Top             =   5520
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
      RecordSource    =   "SalesOrder"
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
      Left            =   4920
      Top             =   5520
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
      RecordSource    =   "SalesOrderLine"
      Caption         =   "SalesOrderLines"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   1320
      Top             =   6000
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
      RecordSource    =   "TempLines"
      Caption         =   "TempLines"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   6720
      Top             =   6120
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
      RecordSource    =   "Customer"
      Caption         =   "Customer"
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
      Caption         =   "Search by Customer Name:"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "SalesForm"
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
'delete all rows in temp table
If Adodc2.recordSet.RecordCount <> 0 Then
    Adodc2.recordSet.MoveFirst
End If
While Not Adodc2.recordSet.EOF
    cursorLocation1 = Adodc2.recordSet.CursorLocation
    If Adodc2.recordSet.Fields("srid") = Adodc1.recordSet.Fields("srid") Then
        Adodc2.recordSet.Delete
    End If
    If cursorLocation1 = Adodc2.recordSet.CursorLocation Then
        Adodc2.recordSet.MoveNext
    End If
Wend
Adodc2.recordSet.Requery
Adodc2.Refresh
Adodc1.recordSet.Delete
End Sub

Private Sub Add_Click()
Me.Hide
SalesReceipt.Show
Adodc1.recordSet.AddNew "OrderDate", Now
Adodc1.recordSet.Update
SalesReceipt.SRID.Text = Adodc1.recordSet.Fields("SRID")

'delete all rows in temp table
While Not SalesReceipt.Adodc1.recordSet.EOF
    SalesReceipt.Adodc1.recordSet.MoveFirst
    If Not SalesReceipt.Adodc1.recordSet.EOF Then
        SalesReceipt.Adodc1.recordSet.Delete
    End If
Wend
SalesReceipt.Adodc1.recordSet.Requery
End Sub

Private Sub Search_Change()
If Search.Text <> "" Then
    If Not search2(Adodc1.recordSet, "CUSTOMERNAME", Search.Text & "*") Then
        MsgBox "Not Found"
        Search.Text = ""
        Adodc1.recordSet.MoveFirst
    End If
End If
End Sub

Private Sub View_Click()
Me.Hide
Load DR_Invoice
DR_Invoice.Show

DR_Invoice.Sections("Section2").Controls.Item("lbl_OrderID").Caption = Adodc1.recordSet.Fields("SRID")

Adodc4.recordSet.MoveFirst
While Not Adodc4.recordSet.EOF
    If Adodc4.recordSet.Fields("CUSTOMERID") = Adodc1.recordSet.Fields("CUSTOMERID") Then
        DR_Invoice.Sections("Section2").Controls.Item("lbl_SoldTo").Caption = Adodc4.recordSet.Fields("CUSTOMERID") & vbCrLf & Adodc4.recordSet.Fields("CUSTOMERNAME") & vbCrLf & Adodc4.recordSet.Fields("Address") & vbCrLf & Adodc4.recordSet.Fields("Phone")
    End If
    Adodc4.recordSet.MoveNext
Wend
Adodc4.recordSet.Requery

DR_Invoice.Sections("Section2").Controls.Item("lbl_Date").Caption = Adodc1.recordSet.Fields("ORDERDATE")
DR_Invoice.Sections("Section5").Controls.Item("Label3").Caption = Adodc1.recordSet.Fields("Discount") & " %"
DR_Invoice.Sections("Section5").Controls.Item("lbl_Total").Caption = Adodc1.recordSet.Fields("TOTALAMT")

Adodc3.recordSet.Requery
Adodc3.Refresh
'delete all rows in temp table
While Not Adodc3.recordSet.EOF
    Adodc3.recordSet.MoveFirst
    If Not Adodc3.recordSet.EOF Then
        Adodc3.recordSet.Delete
    End If
Wend
Adodc3.recordSet.Requery

'copy salesorderlines to temp lines
If Adodc2.recordSet.RecordCount <> 0 Then
    Adodc2.recordSet.MoveFirst
End If
While Not Adodc2.recordSet.EOF
    If Adodc2.recordSet.Fields("SRID") = Adodc1.recordSet.Fields("SRID") Then
        Adodc3.recordSet.AddNew Array("SRID", "PRODUCTCODE", "PRODUCTNAME", "QUANTITY", "SELLINGPRICE", "TOTAL"), Array(Adodc2.recordSet!SRID, Adodc2.recordSet!ProductCode, Adodc2.recordSet!ProductName, Adodc2.recordSet!Quantity, Adodc2.recordSet!SellingPrice, Adodc2.recordSet!total)
    End If
    Adodc2.recordSet.MoveNext
Wend
    
    
Set DR_Invoice.DataSource = Adodc3.recordSet
End Sub
