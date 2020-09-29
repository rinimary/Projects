VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00008080&
   Caption         =   "Client"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4800
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Data to Client/Server"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2685
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP"
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
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 6001
Winsock1.Connect

End Sub

Private Sub Command2_Click()
Winsock1.SendData Text2.Text
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim A As String
Winsock1.GetData A
Text3.Text = A
End Sub

