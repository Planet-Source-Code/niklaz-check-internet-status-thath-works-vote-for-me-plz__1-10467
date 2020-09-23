VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmCheck 
   Caption         =   "[Not checked]"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1470
   ControlBox      =   0   'False
   Icon            =   "frmCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   1470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Red 
      Height          =   615
      Left            =   2520
      Picture         =   "frmCheck.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   2160
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.PictureBox Green 
      Height          =   615
      Left            =   3120
      Picture         =   "frmCheck.frx":0884
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   2160
      Width           =   615
      Visible         =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Status 
      Height          =   480
      Left            =   480
      Picture         =   "frmCheck.frx":0CC6
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Close
Winsock1.Connect "server.napster.com", "8875"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Winsock1_Connect()
Status = Green
Me.Caption = "[Online]"
Winsock1.Close
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Status = Red
Me.Caption = "[Offline]"
End Sub
