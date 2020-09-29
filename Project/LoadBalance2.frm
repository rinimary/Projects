VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLoadBalance2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Load Balancer 2 (Apparels)"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Server 3"
      Height          =   2775
      Index           =   2
      Left            =   6120
      TabIndex        =   17
      Top             =   4920
      Width           =   3015
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send 2"
         Enabled         =   0   'False
         Height          =   735
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtServerIP 
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect 2"
         Enabled         =   0   'False
         Height          =   735
         Index           =   2
         Left            =   480
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
      End
      Begin MSWinsockLib.Winsock WinsockServer 
         Index           =   2
         Left            =   2520
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Server 2"
      Height          =   2775
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   4920
      Width           =   3015
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect 2"
         Height          =   735
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtServerIP 
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send 2"
         Height          =   735
         Index           =   1
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin MSWinsockLib.Winsock WinsockServer 
         Index           =   1
         Left            =   2520
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Server 1"
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   3015
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send 1"
         Height          =   735
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtServerIP 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect 1"
         Height          =   735
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Top             =   1800
         Width           =   2055
      End
      Begin MSWinsockLib.Winsock WinsockServer 
         Index           =   0
         Left            =   2520
         Top             =   2280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SEND"
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LISTEN"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   7
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   270
   End
End
Attribute VB_Name = "frmLoadBalance2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect1_Click()

End Sub

Private Sub cmdSend1_Click()

End Sub



Private Sub cmdConnect_Click(Index As Integer)
    If Index = 0 Then
        WinsockServer(0).RemoteHost = txtServerIP(0)
        WinsockServer(0).RemotePort = 9001
        WinsockServer(0).Connect
        cmdConnect(0).Enabled = False
    End If
    
    If Index = 1 Then
        WinsockServer(1).RemoteHost = txtServerIP(1)
        WinsockServer(1).RemotePort = 9002
        WinsockServer(1).Connect
        cmdConnect(1).Enabled = False
    End If
    
    If Index = 2 Then
        WinsockServer(2).RemoteHost = txtServerIP(2)
        WinsockServer(2).RemotePort = 9003
        WinsockServer(2).Connect
        cmdConnect(2).Enabled = False
    End If
    
End Sub

Private Sub Command1_Click()
Text1.Text = Winsock1.LocalIP
End Sub

Private Sub Command2_Click()
Winsock1.LocalPort = 7002
Winsock1.Listen
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Winsock1.SendData Text2.Text
End Sub

Private Sub Text2_Change()
    Winsock1.SendData Text2
End Sub

Private Sub Text3_Change()
    On Error Resume Next
    
    Dim sqlQuery As String
    
    dbConnect

    sqlQuery = "Select * from tblCluster where category='" & Text3 & "'"

    rs.Open sqlQuery, con, adOpenStatic, adLockOptimistic
        
    txtResult = rs(1)
End Sub

Private Sub txtResult_Change()
    Text2 = ""
    If txtResult = "1" Then
        WinsockServer(0).SendData Text3
    End If
    
    If txtResult = "2" Then
        WinsockServer(1).SendData Text3
    End If
    If txtResult = "3" Then
        WinsockServer(2).SendData Text3
    End If
    
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Winsock1.GetData a
Text3.Text = a
End Sub


Private Sub WinsockServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim a As String
    
    If Index = 0 Then
        WinsockServer(0).GetData a
        Text2 = a
    End If
    
    If Index = 1 Then
        WinsockServer(1).GetData a
        Text2 = a
    End If
    
    If Index = 2 Then
        WinsockServer(2).GetData a
        Text2 = a
    End If
End Sub

