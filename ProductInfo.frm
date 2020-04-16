VERSION 5.00
Begin VB.Form ProductInfo 
   Caption         =   "ProductInfo"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
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
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   3
      Top             =   1800
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
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Save 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Purchase Rate"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Selling Rate"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "ProductInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mode As String
Public str1, str2, str3, str4 As String

Private Sub Cancel_Click()
If Mode = "Edit" Then
    ProductForm.Adodc1.recordSet.AddNew Array("PRODUCTCODE", "PRODUCTNAME", "PURCHASERATE", "SELLINGRATE"), Array(str1, str2, str3, str4)
End If
Unload Me
ProductForm.Show
End Sub

Private Sub Save_Click()
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "Please enter all the values"
    ElseIf Not IsNumeric(Text3.Text) Or Not IsNumeric(Text4.Text) Then
        MsgBox "Please enter numeric values."
    Else
        If Search("product", "PRODUCTCODE", Text1.Text) Then
            MsgBox "ProductCode already exists. Please re-enter a different Code."
            ProductForm.Adodc1.recordSet.MoveFirst
        Else
            ProductForm.Adodc1.recordSet.AddNew Array("PRODUCTCODE", "PRODUCTNAME", "PURCHASERATE", "SELLINGRATE"), Array(Text1.Text, Text2.Text, Text3.Text, Text4.Text)
            Unload Me
            ProductForm.Show
        End If
    End If
End Sub
