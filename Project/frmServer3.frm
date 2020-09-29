VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer13 
   Caption         =   "Server 13 (Shirt Server)"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtReceive 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtSend 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdListen_Click()
    Winsock1.LocalPort = 8003
    Winsock1.Listen
    cmdListen.Enabled = False
End Sub

Private Sub cmdSend_Click()
    Winsock1.SendData txtSend
End Sub

Private Sub Form_Load()
    txtIP = Winsock1.LocalIP
End Sub

Private Sub txtReceive_Change()
    
    On Error Resume Next
    
    txtSend = ""
    Dim sqlQuery As String
    
    dbConnectServer3

    sqlQuery = "Select * from tblServerData where category='" & txtReceive & "'"

    rs3.Open sqlQuery, con3, adOpenStatic, adLockOptimistic
        
    For i = 0 To rs3.RecordCount
        txtSend = txtSend & "@" & rs3(1)
        rs3.MoveNext
    Next i
    
    If i = rs3.RecordCount + 1 Then
        Winsock1.SendData txtSend
    End If
End Sub

Private Sub txtSend_Change()
'    Winsock1.SendData txtSend
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock1.GetData a
    txtReceive = a
End Sub
