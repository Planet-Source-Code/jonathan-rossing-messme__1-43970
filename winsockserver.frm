VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   Icon            =   "winsockserver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock none 
      Left            =   5280
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox UserOnline 
      ForeColor       =   &H80000002&
      Height          =   1425
      Index           =   1
      ItemData        =   "winsockserver.frx":5C12
      Left            =   30
      List            =   "winsockserver.frx":5C14
      TabIndex        =   5
      Top             =   3120
      Width           =   1875
   End
   Begin VB.ListBox UserOnline 
      ForeColor       =   &H80000002&
      Height          =   1425
      Index           =   0
      ItemData        =   "winsockserver.frx":5C16
      Left            =   1920
      List            =   "winsockserver.frx":5C18
      TabIndex        =   4
      Top             =   3120
      Width           =   1875
   End
   Begin MSWinsockLib.Winsock main 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtchat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
   Begin VB.TextBox txtmsg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Send System Message"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2400
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Top             =   3240
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "User online"
      Height          =   285
      Left            =   1950
      TabIndex        =   7
      Top             =   2850
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Reg Users"
      Height          =   285
      Left            =   30
      TabIndex        =   6
      Top             =   2850
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Users As Long
Dim user As String
Dim tmp As String
Dim uonline As Boolean
Dim rad As String

Private Sub Command1_Click()
Dim i As Integer
For i = 0 To 200
If Winsock1(i).State = sckConnected Then Winsock1(i).SendData "Server :" & txtmsg.Text
Next i
txtchat.Text = txtchat.Text & vbCrLf & "Server: " & txtmsg.Text

End Sub

Private Sub Command2_Click()

If Command2.Caption = "Start" Then
main.LocalPort = 806
main.Listen
Command2.Caption = "Stop"
Exit Sub
End If
If Command2.Caption = "Stop" Then
main.Close
Command2.Caption = "Start"
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Form1.Caption = "Server " & none.LocalIP
Open App.Path & "\users.dat" For Input As #2
    Do Until (EOF(2))
        Line Input #2, rad
        UserOnline(1).AddItem rad
    Loop
Close #2
For i = 1 To 200
 Load Winsock1(i)
Next i
End Sub





Private Sub main_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer, j As Integer
For i = 0 To 200
If Winsock1(i).State = sckClosed Then
Users = Users + 1
Winsock1(i).Accept requestID

For j = 0 To 200
If Winsock1(j).State = sckConnected Then Winsock1(j).SendData "NUSER" & Users
Next j

Exit Sub
End If
Next i
main.Close
main.Listen
End Sub

Private Sub Winsock1_Close(Index As Integer)

Users = Users - 1

Dim i As Integer
For i = 0 To 200
If Winsock1(i).State = sckConnected Then Winsock1(i).SendData "DUSER" & Users
Next i

End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Integer
Dim strdata As String
Winsock1(Index).GetData strdata
tmp = Split(strdata, "@,@")(0)
    If tmp = "ULOGIN" Or tmp = "ULOGOUT" Then
        Online strdata, Index
        Exit Sub
    End If
txtchat.Text = txtchat.Text & vbCrLf & strdata
For i = 0 To 200
If Winsock1(i).State = sckConnected And Index <> i Then
Winsock1(i).SendData strdata

End If
Next i

End Sub

Sub Online(Sdata As String, Nr As Integer)
Dim t As Integer
Dim j As Integer
tmp = Split(Sdata, "@,@")(0)
    If tmp = "ULOGIN" Then
        uonline = False
        For t = 0 To UserOnline(1).ListCount
                user = Split(Sdata, "@,@")(1)
            If user = UserOnline(1).List(t) Then
                uonline = True
                UserOnline(0).AddItem (user)
                Winsock1(Nr).SendData "Connection accepted"
                txtchat.Text = txtchat.Text & vbCrLf & Sdata

            Exit Sub
            End If
        Next t
        If uonline = False Then
            Users = Users - 1
            Winsock1(Nr).Close

            txtchat.Text = txtchat.Text & vbCrLf & Sdata
            Exit Sub
        End If
    End If
    
    If tmp = "ULOGOUT" Then
        uonline = False
        For t = 0 To UserOnline(0).ListCount
                user = Split(Sdata, "@,@")(1)
            If user = UserOnline(0).List(t) Then
                uonline = True
                Winsock1(Nr).SendData "Bye Bye"
                UserOnline(0).RemoveItem (t)
                txtchat.Text = txtchat.Text & vbCrLf & Sdata
            Exit Sub
            End If
        Next t
        If uonline = False Then
           txtchat.Text = txtchat.Text & vbCrLf & Sdata
            Exit Sub
        End If
    End If
End Sub

