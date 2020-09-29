VERSION 5.00
Begin VB.Form frmIntro 
   Caption         =   "Clustering"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      Caption         =   "Server 33"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   13
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Server 32"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   12
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Server 23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Server 22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Server 13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Server 12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Load Balancer 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Load Balancer 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Controller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Server 31"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Server 11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Server 21"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Balancer 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   4080
      X2              =   6600
      Y1              =   2760
      Y2              =   3960
   End
   Begin VB.Line Line6 
      X1              =   3840
      X2              =   3840
      Y1              =   2760
      Y2              =   3960
   End
   Begin VB.Line Line5 
      X1              =   3600
      X2              =   1080
      Y1              =   2760
      Y2              =   3840
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   3840
      Y1              =   4920
      Y2              =   6720
   End
   Begin VB.Line Line3 
      X1              =   1320
      X2              =   1320
      Y1              =   4920
      Y2              =   6720
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6600
      Y1              =   4920
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   960
      Y2              =   2280
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    frmSearch.Show
End Sub

Private Sub Command11_Click()
    frmServer22.Show
End Sub

Private Sub Command13_Click()
    frmServer32.Show
End Sub

Private Sub Command2_Click()
    frmLoadBalance1.Show
End Sub

Private Sub Command3_Click()
    frmServer21.Show
End Sub

Private Sub Command4_Click()
    frmServer11.Show
End Sub

Private Sub Command5_Click()
    frmServer31.Show
End Sub

Private Sub Command6_Click()
    frmController.Show

End Sub

Private Sub Command7_Click()
    frmLoadBalance2.Show
End Sub

Private Sub Command8_Click()
    frmLoadBalance3.Show
End Sub

Private Sub Command9_Click()
    frmServer12.Show
End Sub
