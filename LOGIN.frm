VERSION 5.00
Begin VB.Form LOGIN 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LOGIN"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   2280
      TabIndex        =   2
      Top             =   1485
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   1485
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1590
      TabIndex        =   0
      Top             =   450
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1590
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   765
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Attempt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   105
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   0
      Left            =   750
      TabIndex        =   5
      Top             =   495
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   210
      Index           =   1
      Left            =   795
      TabIndex        =   4
      Top             =   795
      Width           =   765
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c_attempt

Private Sub cmdCancel_Click()
    Unload Me

End Sub

Private Sub cmdOk_Click()
c_attempt = c_attempt - 1

If c_attempt < 1 Then
    MsgBox "Sorry all attempts are gone. Application will close now.", vbOKOnly
    Unload Me
Else
    If Text1.Text = "admin" And Text2.Text = "admin" Then 'TODO
        Unload Me
        MainForm.Show
    Else
        MsgBox "Invalid username/password"
        Label7.Caption = c_attempt
        Text1.Text = ""
        Text2.Text = ""
    End If
End If
    
End Sub

Private Sub Form_Load()
c_attempt = 3
Call initializeConnection
End Sub

