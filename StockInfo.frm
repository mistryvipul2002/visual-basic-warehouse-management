VERSION 5.00
Begin VB.Form StockInfo 
   Caption         =   "StockInfo"
   ClientHeight    =   2385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "StockInfo.frx":0000
      Left            =   2280
      List            =   "StockInfo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Quantity"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "StockInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mode As String
Public str1, str2 As String

Private Sub Cancel_Click()
If Mode = "Edit" Then
    StockForm.Adodc1.recordSet.AddNew Array("PRODUCTCODE", "Quantity"), Array(str1, str2)
End If
Unload Me
StockForm.Show
End Sub

Private Sub Save_Click()
    Dim str As String
    If Mode = "Add" Then
        str = Combo1.Text
    Else
        str = Text1.Text
    End If
    
    If str = "" Then
        MsgBox "Please select the correct value"
    ElseIf Text2.Text = "" Then
        MsgBox "Please enter all the values"
    ElseIf Not IsNumeric(Text2.Text) Then
        MsgBox "Please enter numeric values."
    ElseIf Search("stock", "PRODUCTCODE", str) Then
        MsgBox "PRODUCTCODE already exists. Please re-enter a different Code."
        StockForm.Adodc1.recordSet.MoveFirst
    Else
        StockForm.Adodc1.recordSet.AddNew Array("PRODUCTCODE", "Quantity"), Array(str, Text2.Text)
        Unload Me
        StockForm.Show
    End If
End Sub
