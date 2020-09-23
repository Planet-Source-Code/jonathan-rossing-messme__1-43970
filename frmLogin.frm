VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2145
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1267.336
   ScaleMode       =   0  'User
   ScaleWidth      =   3380.205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUid 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1110
      PasswordChar    =   "v"
      TabIndex        =   6
      Top             =   570
      Width           =   2325
   End
   Begin VB.TextBox txtServerip 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1110
      TabIndex        =   4
      Text            =   "xxx.xxx.xxx.xxx"
      Top             =   960
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1110
      TabIndex        =   1
      Text            =   "Master"
      Top             =   180
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   450
      TabIndex        =   2
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2010
      TabIndex        =   3
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   270
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   615
      Width           =   1230
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      Height          =   270
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   1005
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   225
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Form1.Show
   Unload Me
End Sub

Private Sub cmdOK_Click()
        Form1.strSIP = txtServerip
        Form1.N = txtUserName
        Form1.pass = txtUid
        Form1.Sign
        Form1.Show
        Unload Me
        
End Sub

