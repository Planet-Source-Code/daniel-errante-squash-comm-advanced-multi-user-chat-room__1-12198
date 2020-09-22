VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "dancomm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Connect to:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CREATEINI
SAVEINI "USER NAME:", Text1.Text
SAVEINI "SERVER:", Text2.Text
SAVEINI "HOST:", host
SAVEINI "MAX USERS:", musers
un$ = Text1.Text
server$ = Text2.Text
Label3.Caption = "connecting..."
Form2.client.Close
Form2.client.RemoteHost = Text2.Text
Form2.client.RemotePort = 9876
Form2.client.Connect

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
muted = False
op = False
server1 = False
On Error GoTo hell:
THEFILE = App.Path & "\chat.ini"
un$ = READINI("USER NAME:")
server$ = READINI("SERVER:")
host$ = READINI("HOST:")
musers = READINI("MAX USERS:")
lastprt = 8888
Text1.Text = un$
Text2.Text = server$
Load Form2
Form2.ping(0).Close
Form2.ping(0).LocalPort = 9877
Form2.ping(0).Listen

If host$ = "TRUE" Then
    Form2.sock(0).Close
    Form2.sock(0).LocalPort = 9876
    Form2.sock(0).Listen
End If
hell:

End Sub

