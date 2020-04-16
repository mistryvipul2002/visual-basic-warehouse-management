VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Catalogue 
      Caption         =   "Catalogue"
      Begin VB.Menu Customer 
         Caption         =   "Customer"
      End
      Begin VB.Menu Supplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu Product 
         Caption         =   "Product"
      End
      Begin VB.Menu Stock 
         Caption         =   "Stock"
      End
   End
   Begin VB.Menu Transaction 
      Caption         =   "Transaction"
      Begin VB.Menu PurchaseOrder 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu SalesRegister 
         Caption         =   "Sales  Register"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Customer_Click()
Load CustomerForm
CustomerForm.Show
Me.Hide
End Sub

Private Sub Exit_Click()
Unload Me
Call closeConnection
End Sub

Private Sub Product_Click()
Load ProductForm
ProductForm.Show
Me.Hide
End Sub

Private Sub PurchaseOrder_Click()
Load PurchaseForm
PurchaseForm.Show
Me.Hide
End Sub

Private Sub SalesRegister_Click()
Load SalesForm
SalesForm.Show
Me.Hide
End Sub

Private Sub Stock_Click()
Load StockForm
StockForm.Show
Me.Hide
End Sub

Private Sub Supplier_Click()
Load SupplierForm
SupplierForm.Show
Me.Hide
End Sub
