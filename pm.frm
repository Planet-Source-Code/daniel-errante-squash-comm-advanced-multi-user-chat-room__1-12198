VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form4 
   Caption         =   "Private Message To:"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   5850
   Icon            =   "pm.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3795
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3600
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   600
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   2640
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"pm.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   150
      TextRTF         =   $"pm.frx":03DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tim2 As Integer
Private msgpm1 As Integer


Private Sub Form_Load()
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
RichTextBox1.BackColor = READINI("pmbg:")
RichTextBox1.SelStart = Len(RichTextBox1.Text)
'//
RichTextBox2.SelStart = 0
RichTextBox2.SelLength = Len(RichTextBox1.Text)
RichTextBox2.BackColor = READINI("pmbg:")
RichTextBox2.SelStart = Len(RichTextBox1.Text)
RichTextBox2.SelColor = pmclr
RichTextBox2.SelFontSize = pmsze
RichTextBox2.SelBold = pmbld
RichTextBox2.SelFontName = pmfnt
THEFILE = n
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
sock.Close

End Sub

Private Sub Form_Resize()
ResizeForm Me

End Sub

Private Sub RichTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If RichTextBox2.Text = "" Then Exit Sub
 msgs = msgs + 1
                If msgs > 2 Then
        RichTextBox2.Locked = True
            Timer3.Enabled = True
            msgto "You cannot talk for 10 seconds (flood)."
            Timer2.Enabled = False
            Exit Sub
        End If
    If sock.State = 7 Then
        sock.SendData "pmg " & un$ & ": " & RichTextBox2.Text
        DoEvents
        msgto un$ & ": " & RichTextBox2.Text, vbBlue
    End If
    RichTextBox2.Text = ""
End If
End Sub

Private Sub sock_Close()
MsgBox "Disconnected from " & Me.Caption, vbExclamation
Unload Me
End Sub

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
sock.Close
sock.Accept requestID
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
sock.GetData strdata
cmd$ = Left(strdata, 4)
rest$ = Right(strdata, Len(strdata) - 4)
If cmd$ = "pmg " Then msgto rest$

End Sub


Private Function msgto(message As String, Optional clr As String = vbBlack, Optional bld As Boolean = False, Optional sze As Long = 10)
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelColor = pmclr
RichTextBox1.SelFontSize = pmsze
RichTextBox1.SelBold = pmbld
RichTextBox1.SelFontName = pmfnt
RichTextBox1.SelText = message & vbCrLf
RichTextBox1.SelStart = Len(RichTextBox1.Text)

End Function

Private Sub Timer2_Timer()
msgpm1 = 0
End Sub

Private Sub Timer3_Timer()
tim2 = tim2 + 1
RichTextBox2.Locked = True
If tim2 = 10 Then
Timer3.Enabled = False
Timer2.Enabled = True
msgto "You can talk again now."
    RichTextBox2.Locked = False
    tim2 = 0
    msgspm1 = 0
End If
End Sub
