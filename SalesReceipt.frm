VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form SalesReceipt 
   Caption         =   "Sales Receipt"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextDiscount 
      Height          =   285
      Left            =   1920
      TabIndex        =   25
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "CustomerID"
      DataSource      =   "Adodc4"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TextCustomerName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8880
      TabIndex        =   21
      Top             =   6000
      Width           =   4455
   End
   Begin VB.ComboBox ComboCustomerID 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox SRID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Text            =   "####"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11640
      TabIndex        =   16
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton DeleteRow 
      Caption         =   "Delete Row"
      Height          =   375
      Left            =   11880
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      DataField       =   "ProductCode"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   12600
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "ProductCode"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox ComboProductCode 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TextStock 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9120
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox TextSellingPrice 
      Height          =   285
      Left            =   7560
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TextQuantity 
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox TextProductName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SalesReceipt.frx":0000
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   7011
      _Version        =   393216
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
      Left            =   240
      Top             =   8280
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   4080
      Top             =   8280
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
      RecordSource    =   "Product"
      Caption         =   "Product"
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
      Left            =   9000
      Top             =   8280
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
      Caption         =   "Stock"
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
      Left            =   240
      Top             =   8760
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
   Begin VB.Label LabelTotal 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label LabelDiscount 
      Caption         =   "Discount %"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Customer Name"
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label LabelCustomerId 
      Caption         =   "Customer ID"
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label LabelSRID 
      Caption         =   "SRID"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Stock"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Selling Price"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Quantity 
      Caption         =   "Qty"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label ProductName 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label LabelProductCode 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "SalesReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
SalesForm.Adodc1.recordSet.Delete
Unload Me
End Sub

Private Sub Add_Click()
If TextQuantity.Text = "" Then
    MsgBox "Please fill all values"
    Exit Sub
ElseIf Not IsNumeric(TextQuantity.Text) Then
    MsgBox "Please fill numeric values"
    Exit Sub
ElseIf Val(TextQuantity.Text) > Val(TextStock.Text) Then
    MsgBox "Please reduce the quantity. You are placing the order more than available stock"
    Exit Sub
End If

Adodc1.recordSet.Requery
While Not Adodc1.recordSet.EOF
    If ComboProductCode.Text = Adodc1.recordSet.Fields("PRODUCTCODE") Then
        newQty = Adodc1.recordSet.Fields("QUANTITY") + TextQuantity.Text
        If newQty > Val(TextStock.Text) Then
            MsgBox "Please reduce the quantity. You are placing the order more than available stock"
            Exit Sub
        Else
            Adodc1.recordSet.Update Array("SRID", "PRODUCTCODE", "PRODUCTNAME", "QUANTITY", "SELLINGPRICE", "TOTAL"), Array(SRID.Text, ComboProductCode.Text, TextProductName.Text, newQty, TextSellingPrice.Text, newQty * TextSellingPrice.Text)
            Call updateTotal
            Exit Sub
        End If
    End If
    Adodc1.recordSet.MoveNext
Wend
Adodc1.recordSet.AddNew Array("SRID", "PRODUCTCODE", "PRODUCTNAME", "QUANTITY", "SELLINGPRICE", "TOTAL"), Array(SRID.Text, ComboProductCode.Text, TextProductName.Text, TextQuantity.Text, TextSellingPrice.Text, TextQuantity.Text * TextSellingPrice.Text)
Call updateTotal
End Sub

Private Sub ComboCustomerID_Click()
Adodc4.recordSet.MoveFirst
While Not Adodc4.recordSet.EOF
    If ComboCustomerID.Text = Adodc4.recordSet.Fields("CUSTOMERID") Then
        TextCustomerName.Text = Adodc4.recordSet.Fields("CUSTOMERNAME")
    End If
    Adodc4.recordSet.MoveNext
Wend
End Sub

Private Sub ComboProductCode_Click()
Adodc2.recordSet.MoveFirst
While Not Adodc2.recordSet.EOF
    If ComboProductCode.Text = Adodc2.recordSet.Fields("PRODUCTCODE") Then
        TextProductName.Text = Adodc2.recordSet.Fields("PRODUCTNAME")
        TextSellingPrice.Text = Adodc2.recordSet.Fields("SELLINGRATE")
    End If
    Adodc2.recordSet.MoveNext
Wend
Adodc3.recordSet.MoveFirst
While Not Adodc3.recordSet.EOF
    If ComboProductCode.Text = Adodc3.recordSet.Fields("PRODUCTCODE") Then
        TextStock.Text = Adodc3.recordSet.Fields("QUANTITY")
    End If
    Adodc3.recordSet.MoveNext
Wend
End Sub

Private Sub DeleteRow_Click()
If Not Adodc1.recordSet.EOF And Not Adodc1.recordSet.BOF Then
    Adodc1.recordSet.Delete
End If
Call updateTotal
End Sub

Private Sub Form_Load()
Adodc1.Refresh

'populate productCode combo box
ComboProductCode.Clear
Adodc3.recordSet.MoveFirst
While Not Adodc3.recordSet.EOF
    If Val(Adodc3.recordSet.Fields("QUANTITY")) > 0 Then
        ComboProductCode.AddItem Adodc3.recordSet.Fields("PRODUCTCODE")
    End If
    Adodc3.recordSet.MoveNext
Wend
ComboProductCode.ListIndex = 0

'populate customerId combo box
ComboCustomerID.Clear
Adodc4.recordSet.MoveFirst
While Not Adodc4.recordSet.EOF
    ComboCustomerID.AddItem Adodc4.recordSet.Fields("CUSTOMERID")
    Adodc4.recordSet.MoveNext
Wend
ComboCustomerID.ListIndex = 0

End Sub

Private Sub Save_Click()
If Adodc1.recordSet.RecordCount = 0 Then
    MsgBox "No products selected. Please add some products."
    Exit Sub
End If

'Add record to salesOrder table
SalesForm.Adodc1.recordSet.Update "ORDERDATE", Now
SalesForm.Adodc1.recordSet.Update "customerID", ComboCustomerID.Text
SalesForm.Adodc1.recordSet.Update "CustomerName", TextCustomerName.Text
SalesForm.Adodc1.recordSet.Update "Discount", Val(TextDiscount.Text)
SalesForm.Adodc1.recordSet.Update "TotalAmt", Val(LabelTotal.Caption)

'Copy records from tempLines to salesOrderLines table
If Adodc1.recordSet.RecordCount <> 0 Then
    Adodc1.recordSet.MoveFirst
    While Not Adodc1.recordSet.EOF
        SalesForm.Adodc2.recordSet.AddNew Array("SRID", "PRODUCTCODE", "PRODUCTNAME", "QUANTITY", "SELLINGPRICE", "TOTAL"), Array(Adodc1.recordSet!SRID, Adodc1.recordSet!ProductCode, Adodc1.recordSet!ProductName, Adodc1.recordSet!Quantity, Adodc1.recordSet!SELLINGPRICE, Adodc1.recordSet!total)
        
        'reduce stock
        If Adodc3.recordSet.RecordCount <> 0 Then
            Adodc3.recordSet.MoveFirst
            While Not Adodc3.recordSet.EOF
                If Adodc1.recordSet!ProductCode = Adodc3.recordSet.Fields("PRODUCTCODE") Then
                    Adodc3.recordSet.Update "Quantity", Val(Adodc3.recordSet.Fields("QUANTITY")) - Adodc1.recordSet!Quantity
                End If
                Adodc3.recordSet.MoveNext
            Wend
        End If
        
         'update purchase price for products
        If Adodc2.recordSet.RecordCount <> 0 Then
            Adodc2.recordSet.MoveFirst
            While Not Adodc2.recordSet.EOF
                If Adodc1.recordSet!ProductCode = Adodc2.recordSet.Fields("PRODUCTCODE") Then
                    Adodc2.recordSet.Update "SELLINGRATE", Adodc1.recordSet!SELLINGPRICE
                End If
                Adodc2.recordSet.MoveNext
            Wend
        End If
        Adodc1.recordSet.MoveNext
    Wend
End If
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SalesForm.Show
End Sub

Private Sub updateTotal()
Dim total As Double
total = 0#
If Adodc1.recordSet.RecordCount <> 0 Then
    Adodc1.recordSet.MoveFirst
    While Not Adodc1.recordSet.EOF
        total = total + Val(Adodc1.recordSet.Fields("Total"))
    Adodc1.recordSet.MoveNext
    Wend
End If
total = total - (total * (Val(TextDiscount.Text) / 100#))
LabelTotal.Caption = total
If Adodc1.recordSet.RecordCount <> 0 Then
    Adodc1.recordSet.MoveFirst
End If
End Sub

Private Sub TextDiscount_Change()
Call updateTotal
End Sub
