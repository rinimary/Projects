VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmController 
   Caption         =   "Controller"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   5400
      TabIndex        =   10
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   2760
      TabIndex        =   9
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect to Load Balancer 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect to Load Balancer 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox LBIP3 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   5400
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox LBIP2 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   2760
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox LBIP1 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   2880
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   120
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect to Load Balancer 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtRecv 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   2760
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   5400
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListen_Click()
    Winsock1.LocalPort = 6000
    Winsock1.Listen
    cmdListen.Enabled = False
End Sub

Private Sub cmdConnect_Click()
    Winsock2.RemotePort = 7001
    Winsock2.RemoteHost = LBIP1
    Winsock2.Connect
    cmdConnect.Enabled = False
End Sub

Private Sub Command1_Click()
    Winsock3.RemotePort = 7002
    Winsock3.RemoteHost = LBIP2
    Winsock3.Connect
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()
    Winsock4.RemotePort = 7003
    Winsock4.RemoteHost = LBIP3
    Winsock4.Connect
    Command2.Enabled = False
End Sub

Private Sub txtRecv_Change()
    'On Error Resume Next
    
    Dim sqlQuery As String
    
    dbController
                                
    sqlQuery = "Select * from tblController where category='" & txtRecv & "'"

    rs4.Open sqlQuery, con4, adOpenStatic, adLockOptimistic
        
        
    If Val(rs4(4)) = 1 Then
        'LBIP1 = rs4(2)
        Winsock2.SendData txtRecv
    End If
    
    If Val(rs4(4)) = 2 Then
        'LBIP2 = rs4(2)
        Winsock3.SendData txtRecv
    End If
    
    If Val(rs4(4)) = 3 Then
        'LBIP3 = rs4(2)
        Winsock4.SendData txtRecv
    End If
    
    
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock1.GetData a
    txtRecv = a
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock2.GetData a
    
    Text1 = a
    Winsock1.SendData a
    
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock3.GetData a
    
    Text2 = a
    Winsock1.SendData a
End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock4.GetData a
    
    Text3 = a
    Winsock1.SendData a
End Sub
