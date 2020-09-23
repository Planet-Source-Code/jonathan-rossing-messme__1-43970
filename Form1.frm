VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MessMe"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8760
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":5C12
   ScaleHeight     =   5790
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   5070
      TabIndex        =   4
      Top             =   3960
      Width           =   1755
      Visible         =   0   'False
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   8
         Left            =   360
         Picture         =   "Form1.frx":AB40E
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   13
         Tag             =   ";)"
         Top             =   150
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   30
         Picture         =   "Form1.frx":AB571
         ScaleHeight     =   240
         ScaleWidth      =   315
         TabIndex        =   12
         Tag             =   ":("
         Top             =   150
         Width           =   315
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   6
         Left            =   630
         Picture         =   "Form1.frx":ABD2C
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   11
         Tag             =   ":p"
         Top             =   150
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   5
         Left            =   630
         Picture         =   "Form1.frx":ABE09
         ScaleHeight     =   225
         ScaleWidth      =   510
         TabIndex        =   10
         Tag             =   ":*"
         Top             =   420
         Width           =   510
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   4
         Left            =   360
         Picture         =   "Form1.frx":AC4E4
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   9
         Tag             =   ":)"
         Top             =   420
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   1170
         Picture         =   "Form1.frx":AD912
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   8
         Tag             =   ":x"
         Top             =   420
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   1440
         Picture         =   "Form1.frx":ADA30
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   7
         Tag             =   ">=)"
         Top             =   150
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   900
         Picture         =   "Form1.frx":ADCCD
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   6
         Tag             =   ":-s"
         Top             =   150
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   1170
         Picture         =   "Form1.frx":AE00C
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   5
         Tag             =   ":D"
         Top             =   150
         Width           =   225
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   9
         Left            =   90
         Picture         =   "Form1.frx":AE0B6
         ScaleHeight     =   330
         ScaleWidth      =   225
         TabIndex        =   14
         Tag             =   ":?"
         Top             =   300
         Width           =   225
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9270
      Top             =   5250
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "Form1.frx":AE587
      Top             =   6480
      Width           =   6255
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "&s"
      DownPicture     =   "Form1.frx":AE7CC
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   6870
      MaskColor       =   &H000080FF&
      Picture         =   "Form1.frx":B3D10
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Alt+s"
      Top             =   3090
      UseMaskColor    =   -1  'True
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      Height          =   1425
      HideSelection   =   0   'False
      Left            =   5100
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4140
      Width           =   3465
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      DownPicture     =   "Form1.frx":B9254
      Height          =   525
      Left            =   5070
      MaskColor       =   &H000080FF&
      Picture         =   "Form1.frx":BC029
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3510
      Width           =   1755
   End
   Begin SHDocVwCtl.WebBrowser W 
      Height          =   4665
      Left            =   330
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   8229
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Users:"
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
      Left            =   90
      TabIndex        =   19
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   870
      TabIndex        =   18
      Top             =   90
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      Height          =   225
      Left            =   5850
      TabIndex        =   16
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nr Chr left:"
      Height          =   225
      Left            =   5040
      TabIndex        =   17
      Top             =   3240
      Width           =   825
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuSign 
         Caption         =   "Sign in"
      End
      Begin VB.Menu mnuIP 
         Caption         =   "Change Server IP"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mess As String
Dim MsgSend As String
Dim user As String
Dim smile(11) As String
Dim name1 As String
Public strSIP As String
Public N As String
Public pass As String
Dim ya As Integer

Private Sub Command1_Click()
Frame2.Visible = False
Send
Text1.SetFocus
End Sub
Sub msg(s As String)

Dim newline As String

    If Winsock1.State = sckConnected Then
        Winsock1.SendData name1 & s
    End If
    
    
    Open App.Path & "\temp.html" For Append As #1
    Print #1, name1
    Print #1, s
    Close #1
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
If Frame2.Visible = True Then
Frame2.Visible = False
Else
Frame2.Visible = True
End If
End Sub

Private Sub Form_Click()
Frame2.Visible = False
End Sub

Private Sub Form_Load()



DelFile
MakeFile
W.Navigate "file://" & App.Path & "\temp.html"
smile(0) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/1.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(1) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/2.gif" & Chr(34) & " width=" & Chr(34) & "21" & Chr(34) & " height=" & Chr(34) & "16" & Chr(34) & ">"
smile(2) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/3.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(3) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/4.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(4) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/5.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(5) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/6.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "22" & Chr(34) & ">"
smile(6) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/7.gif" & Chr(34) & " width=" & Chr(34) & "34" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(7) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/8.gif" & Chr(34) & " width=" & Chr(34) & "65" & Chr(34) & " height=" & Chr(34) & "23" & Chr(34) & ">"
smile(8) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/9.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
smile(9) = "<img border=" & Chr(34) & "0" & Chr(34) & " src=" & Chr(34) & "http://hem.passagen.se/master_kane/10.gif" & Chr(34) & " width=" & Chr(34) & "15" & Chr(34) & " height=" & Chr(34) & "15" & Chr(34) & ">"
Form1.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Winsock1.State = sckConnected Then
Winsock1.SendData ("ULOGOUT@,@" & N & "@.@" & pass)
DoEvents
End If
Winsock1.Close
DelFile
Unload frmLogin
Unload frmSplash
End Sub

Sub Send()
If Text1 = "" Or Text1 = " " Or Text1 = "  " Then Exit Sub
    mess = Replace(Text1, ":)", smile(0))
    mess = Replace(mess, ":(", smile(1))
    mess = Replace(mess, ";)", smile(2))
    mess = Replace(mess, ":D", smile(3))
    mess = Replace(mess, ":p", smile(4))
    mess = Replace(mess, ":?", smile(5))
    mess = Replace(mess, ":*", smile(6))
    mess = Replace(mess, ">=)", smile(7))
    mess = Replace(mess, ":x", smile(8))
    mess = Replace(mess, ":-s", smile(9))
    mess = Replace(mess, "javascript", "java´script")
    mess = Replace(mess, "</script>", "=) ajabaja")
    mess = Replace(mess, "a href=", "a target=" & Chr(34) & "_blank" & Chr(34) & "href=")
    mess = Replace(mess, Chr$(13), "<br>")
    mess = Replace(mess, Chr$(10), "<br>")
    MsgSend = "<b><font color=" & "green" & ">" & mess & "</font></b></p>"
    msg (mess)
    Text1 = ""
   W.Navigate "file://" & App.Path & "\temp.html"

   
    
End Sub

Private Sub mnuIP_Click()
strSIP = InputBox("Server IP: ", "Messenger!", "195.149.165.56")

End Sub

Private Sub mnuSign_Click()
If mnuSign.Caption = "Sign out" Then
Sign
Else
frmLogin.Show
End If

End Sub

Private Sub P_Click(Index As Integer)
Text1 = Text1 + P(Index).Tag
Frame2.Visible = False
Text1.SetFocus
End Sub

Private Sub Picture1_Click(Index As Integer)

End Sub

Private Sub P_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
P(Index).ToolTipText = P(Index).Tag

End Sub

Private Sub Text1_Change()
Dim L As String
L = Text1
Label5 = 200 - Len(L)
End Sub

Private Sub Text1_Click()
Frame2.Visible = False
End Sub

Sub MakeFile()
Attribute MakeFile.VB_Description = "Make the temp file"
    Open App.Path & "\temp.html" For Append As #1
    Print #1, Text2
    Close #1
End Sub

Sub DelFile()
On Error Resume Next
Kill App.Path & "\temp.html"
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData ("ULOGIN@,@" & N & "@.@" & pass)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
Winsock1.GetData strdata
If Left(strdata, 5) = "DUSER" Then
Label4.Caption = Right(strdata, Len(strdata) - 5)
Exit Sub
End If
If Left(strdata, 5) = "NUSER" Then
Label4.Caption = Right(strdata, Len(strdata) - 5)
Exit Sub
End If

W.Navigate "file://" & App.Path & "\temp.html"
    Open App.Path & "\temp.html" For Append As #1
    Print #1, strdata
    Close #1
    
End Sub
Sub Sign()
user = frmLogin.txtUserName
name1 = "<p><b><font color=" & "blue" & "size=" & "4" & ">" & user & " says: </font></b><br>"
Form1.Caption = user

If mnuSign.Caption = "Sign in" Then

If strSIP = "" Then Exit Sub
Winsock1.RemoteHost = strSIP
Winsock1.RemotePort = 806
Winsock1.Connect
Text1.Enabled = True
Command1.Enabled = True
mnuSign.Caption = "Sign out"
Exit Sub
End If
If mnuSign.Caption = "Sign out" Then
Winsock1.Close
Text1.Enabled = False
Command1.Enabled = False
mnuSign.Caption = "Sign in"

Exit Sub
End If
End Sub

