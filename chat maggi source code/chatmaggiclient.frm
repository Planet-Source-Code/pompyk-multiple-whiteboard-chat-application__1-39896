VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "chat maggi ---- client"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "chatmaggiclient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7920
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   7200
      ScaleHeight     =   4635
      ScaleWidth      =   4515
      TabIndex        =   24
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton clear 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9840
      TabIndex        =   23
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox receiveddata 
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   7560
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   10800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   10320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   10320
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   9840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   4
      Left            =   9840
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H0000C0C0&
      Height          =   375
      Index           =   5
      Left            =   9360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00004080&
      Height          =   375
      Index           =   6
      Left            =   9360
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00404000&
      Height          =   375
      Index           =   7
      Left            =   10800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   8
      Left            =   8880
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox backbaby 
      BackColor       =   &H00FFC0FF&
      Height          =   375
      Index           =   9
      Left            =   8880
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   960
      Width           =   375
   End
   Begin VB.ComboBox width1 
      Height          =   315
      ItemData        =   "chatmaggiclient.frx":030A
      Left            =   9000
      List            =   "chatmaggiclient.frx":0323
      TabIndex        =   11
      Top             =   6480
      Width           =   660
   End
   Begin VB.TextBox txtcommon 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3240
      Width           =   6975
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "SEND.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtchat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Width           =   5055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "CLose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmddisconnect 
      Caption         =   "DIsconnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "COnnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame111 
      Caption         =   "SERVER SETTINGS.........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdlocal 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   29
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   28
         Text            =   "2020"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Text            =   "1010"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "DRAWING PORT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "IP :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "CHAT PORT:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "colours....."
      Height          =   1335
      Left            =   8400
      TabIndex        =   25
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PEN SIZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   26
      Top             =   6480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'author: somdutt ganguly
'email: gangulysomdutt@yahoo.com
'address: no 6,chandrodaya apt,bhaikaka nagar
'thaltej, ahmedabad, gujarat, india - 380059
'year: 2001-2002
'status: TY BCA from CPICA College - Gujarat university
'note: plz don't modify the source code...and pretend that
'u r the author since this is not a good practice for
'any programmer.....i appreciate if u make this program
'better.
'i appreciate feed backs...thx
Dim xx As Long
Dim yy As Long
Const length = 10
Dim drawcolor As Long
Private Sub backbaby_Click(Index As Integer)
backbaby(Index).BorderStyle = 0
For i = 0 To backbaby.UBound
If i <> Index Then
backbaby(i).BorderStyle = 1
End If
Next i
drawcolor = backbaby(Index).BackColor

End Sub

Private Sub clear_Click()
Picture1.Cls

End Sub

Private Sub cmdconnect_Click()
On Error GoTo x
Winsock1.Close
Winsock1.RemoteHost = Text2.Text
Winsock1.RemotePort = Text1.Text
Winsock1.Connect
Winsock2.Close
Winsock2.RemoteHost = Text2.Text
Winsock2.RemotePort = Text3.Text
Winsock2.Connect

Exit Sub
x:
MsgBox Err.Description
Winsock1.Close

End Sub

Private Sub cmddisconnect_Click()
End
End Sub

Private Sub cmdlocal_Click()
Text2.Text = Winsock1.LocalIP
End Sub

Private Sub cmdsend_Click()
Dim messagebaby As String
On Error GoTo x
messagebaby = Winsock1.LocalHostName & " : " & _
txtchat.Text & vbCrLf
Winsock1.SendData messagebaby
txtchat.Text = ""
txtchat.SetFocus
Exit Sub
x:
MsgBox Err.Description
Winsock1.Close


End Sub

Private Sub Form_Load()
drawcolor = 0
Picture1.DrawWidth = 1
End Sub

Private Sub width1_Click()
Picture1.DrawWidth = Val(width1.Text)

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data1 As String
'Dim message1 As String
Winsock1.GetData data1, vbString
txtcommon.SelStart = Len(txtchat.Text)
txtcommon.SelText = data1

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo x
Picture1.Line (x, y)-(x, y), drawcolor
xx = x
yy = y
bhajoo sFormatSend(x) & sFormatSend(y) & sFormatSend(x) & sFormatSend(y) & sFormatSend(drawcolor) & sFormatSend(Picture1.DrawWidth)
x:
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo x
If Button = 1 Then
Picture1.Line (xx, yy)-(x, y), drawcolor
bhajoo sFormatSend(xx) & sFormatSend(yy) & sFormatSend(x) & sFormatSend(y) & sFormatSend(drawcolor) & sFormatSend(Picture1.DrawWidth)
End If
xx = x
yy = y
x:
End Sub

Public Function sFormatSend(vData) As String
'Format data to send.
sFormatSend = Format(vData, String(length, "0"))

If Len(sFormatSend) = length + 1 Then
    sFormatSend = Format(vData, String(length - 1, "0"))
End If
End Function

Public Function sParam(vsData As String, viNum As Integer) As String

sParam = Mid(vsData, length * (viNum - 1) + 1, length)
End Function


Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim drawme As String
Winsock2.GetData drawme
receiveddata.Text = drawme

 Picture1.Line (sParam(drawme, 1), sParam(drawme, 2))-(sParam(drawme, 3), sParam(drawme, 4)), sParam(drawme, 5)
 Picture1.DrawWidth = sParam(drawme, 6)
End Sub
Public Sub bhajoo(dataman As String)
Winsock2.SendData dataman
End Sub
