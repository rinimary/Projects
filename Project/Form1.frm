VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "SERVER"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "GET IP"
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEND"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LISTEN"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.LocalPort = 6000
Winsock1.Listen
End Sub

Private Sub Command2_Click()
Winsock1.SendData Text1.Text
End Sub

Private Sub Command3_Click()
Text3.Text = Winsock1.LocalIP
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim A As String
Winsock1.GetData A
Text2.Text = A
End Sub


