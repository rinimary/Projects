VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00808000&
   Caption         =   "User Search"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   6000
      TabIndex        =   8
      Top             =   360
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Data to server"
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   735
      Left            =   720
      TabIndex        =   3
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2280
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
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Products"
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
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP"
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
      Left            =   720
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = 6000
Winsock1.Connect
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Winsock1.SendData Text2.Text
End Sub



Private Sub Text3_Change()
On Error Resume Next

 Dim strTest As String
   Dim strArray() As String
   Dim intCount As Integer
   
   strTest = Mid(Text3, 2) ' "Fred & Wilma & Barney & Betty"
   strArray = Split(strTest, "@")
   
   For intCount = LBound(strArray) To UBound(strArray)
      'Debug.Print Trim(strArray(intCount))
      List1.AddItem Trim(strArray(intCount))
   Next
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim a As String
Winsock1.GetData a
Text3.Text = a
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    Winsock2.Close
    Winsock2.Accept requestID
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim b As String
    Winsock2.GetData b
    Text3 = b
End Sub
