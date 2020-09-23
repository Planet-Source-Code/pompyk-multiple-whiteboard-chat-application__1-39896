VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "chat maggi ----------------- SERVER"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "chatmaggi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   3120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      Height          =   5175
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1560
      Width           =   3975
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
      Height          =   5175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1560
      Width           =   7455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2640
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DISCONNECT"
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
      Left            =   8640
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdlisten 
      Caption         =   "LISTEN"
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
      Left            =   8640
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Text4 
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
         Left            =   5640
         TabIndex        =   9
         Text            =   "2020"
         Top             =   240
         Width           =   1095
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
         Left            =   2040
         TabIndex        =   1
         Text            =   "1010"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DRAW LOCAL PORT:"
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
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CHAT LOCAL PORT:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Left            =   8160
      TabIndex        =   10
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LOG DETAILS:----"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   11760
      Y1              =   1080
      Y2              =   1080
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

Option Explicit
Private noofsockets As Integer
Private noofsocketsdraw As Integer


Private Sub cmdlisten_Click()
Dim localp As Long
On Error GoTo x
localp = Text1.Text
Winsock1(0).LocalPort = localp
Winsock1(0).Listen
localp = Winsock1(0).LocalPort
Text2 = Text2.Text & "listening on port no " & localp & vbCrLf
Winsock2(0).LocalPort = Text4.Text
Winsock2(0).Listen
Text3 = Text3.Text & "listening on port no " & Text4.Text & vbCrLf

Exit Sub

x:
MsgBox Err.Description
Text2 = Text2.Text & "error occured..." & Err.Description & vbCrLf
Text3 = Text3.Text & "error occured " & Err.Description & vbCrLf

Winsock1(0).Close
Winsock2(0).Close

End Sub


Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Label3.Caption = Winsock1(0).LocalIP

End Sub

Private Sub Winsock1_Close(Index As Integer)
Text2 = Text2.Text & "connection closed" & vbCrLf
Winsock1(Index).Close
Unload Winsock1(Index)
Text2 = Text2.Text & "sockets closed by server" & vbCrLf

End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
noofsockets = noofsockets + 1
Load Winsock1(noofsockets)
Winsock1(noofsockets).LocalPort = 0
Winsock1(noofsockets).Accept requestID
Text2 = Text2.Text & " connection received from " & Winsock1(noofsockets).RemoteHostIP & vbCrLf


End If

End Sub


Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String
Dim data1 As String
Dim i As Integer

Winsock1(Index).GetData data, vbString
Text2 = Text2.Text & "got " & bytesTotal & _
" bytes from socket " & Index

On Error Resume Next
For i = 1 To noofsockets
Winsock1(i).SendData data

DoEvents
Next

End Sub


Private Sub Winsock2_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
noofsocketsdraw = noofsocketsdraw + 1
Load Winsock2(noofsocketsdraw)
Winsock2(noofsocketsdraw).LocalPort = 0
Winsock2(noofsocketsdraw).Accept requestID

Text3 = Text3.Text & " connection received from " & Winsock2(noofsocketsdraw).RemoteHostIP & vbCrLf

End If

End Sub


Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String
Dim data1 As String
Dim i As Integer

Winsock2(Index).GetData data, vbString
Text3 = Text3.Text & "got " & bytesTotal & _
" bytes from socket " & Index


On Error Resume Next
For i = 1 To noofsocketsdraw
Winsock2(i).SendData data

DoEvents
Next

End Sub


Private Sub Winsock2_Close(Index As Integer)

Winsock2(Index).Close
Unload Winsock2(Index)

End Sub
