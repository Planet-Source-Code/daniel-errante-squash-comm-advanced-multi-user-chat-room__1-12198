VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "SqUaShComm"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   8040
   Icon            =   "dancomm2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5250
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6720
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   4560
   End
   Begin MSWinsockLib.Winsock ping2 
      Left            =   5760
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ping 
      Index           =   0
      Left            =   5400
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock client 
      Left            =   6840
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sock 
      Index           =   0
      Left            =   6480
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   6600
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393217
      MultiLine       =   0   'False
      MaxLength       =   150
      TextRTF         =   $"dancomm2.frx":030A
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
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8070
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"dancomm2.frx":03DE
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
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopt 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnufilebar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnudisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tme As Long

Private Sub client_Close()
On Error GoTo hell:
ext = True
Unload Me
ext = False
Load Form1
Form1.Show
MsgBox "Disconnected from server.", vbExclamation
Unload Form4
Unload Form5
Unload Form6
Unload frmAbout
op = False
hell:
End Sub

Private Sub client_Connect()
On Error GoTo hell:
send "add " & un$
Unload Form1
Form2.Show
Form2.Caption = "SqUaShComm (" & un$ & ")"
hell:

End Sub

Private Sub client_DataArrival(ByVal bytesTotal As Long)
DoEvents
Dim strdata As String
client.GetData strdata
DoEvents
cmd$ = Left(strdata, 4)
rest$ = Right(strdata, Len(strdata) - 4)
If cmd$ = "msg " Then
rest$ = validatemsg(rest$)
    For d = 1 To Len(rest$)
        m$ = Mid(rest$, d, 1)
            If m$ = ":" Then
                un3$ = j$
            Else
            j$ = j$ & m$
            End If
    Next d
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
For i = 0 To READINI("#ignore:")
    If LCase(un3$) = LCase(READINI("ignore " & i & ":")) Then
        Exit Sub
    End If
Next i
THEFILE = n
msg rest$
End If
If cmd$ = "con " Then
'// parse "con SqUaSh,127.0.0.1,8889"
For i = 1 To Len(rest$)
    m$ = Mid(rest$, i, 1)
        If m$ = "," Then
                un2$ = w$
                w$ = ""
           For d = i + 1 To Len(rest$)
                t$ = Mid(rest$, d, 1)
                    If t$ = "," Then
                        sv$ = j$
                        j$ = ""
                            For h = d + 1 To Len(rest$)
                                r$ = Mid(rest$, h, 1)
                                    If r$ = "." Then
                                        prt$ = e$
                                        e$ = ""
                                        Dim frm5 As New Form5
                                        Set frm5 = New Form5
                                        Load frm5
                                        frm5.Caption = un2$
                                        frm5.client.RemotePort = prt$
                                        frm5.client.RemoteHost = sv$
                                        frm5.client.Connect
                                        If pms = "0" Then
                                            DoEvents
                                            frm5.client.SendData "pmg " & un$ & ": I am not accepting PMs right now."
                                            DoEvents
                                            frm5.client.Close
                                            Exit Sub
                                        End If
                                        frm5.Show
                                        frm5.RichTextBox2.SetFocus
                                    Else
                                        e$ = e$ & r$
                                    End If
            
                            Next h
                    Else
                        j$ = j$ & t$
                    End If
            
           Next d
        Else
        w$ = w$ & m$
        End If
Next i
End If
If cmd$ = "dis " Then msg rest$
If cmd$ = "png " Then
ping2.RemoteHost = rest$
ping2.RemotePort = 9877
ping2.Connect
End If
If cmd$ = "lst " Then
    If hosting = True Then Exit Sub
    List1.Clear
    For i = 1 To Len(rest$)
        t$ = Mid(rest$, i, 1)
            If t$ = "," Then
                If j$ <> "" And j$ <> "," Then List1.AddItem j$
                j$ = ""
            Else
            j$ = j$ & t$
            End If
    Next i
End If
If cmd$ = "!/n " Then
    send "act * " & un$ & " is now known as " & rest$ & "."
    un$ = rest$
    Form2.Caption = "SqUaShComm (" & un$ & ")"
THEFILE = App.Path & "\chat.ini"
    CREATEINI
SAVEINI "USER NAME:", un$
SAVEINI "SERVER:", server$
SAVEINI "HOST:", host
SAVEINI "MAX USERS:", musers
End If
If cmd$ = "!-n " Then
        msg rest$
End If
If cmd$ = "!+o " Then
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                If onick$ = un$ Then
                    op = True
                End If
                msg opnick$ & " opped " & onick$ & "."
            Else
            j$ = j$ & m$
            End If
    Next i
End If
If cmd$ = "!-o " Then
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                If onick$ = un$ Then
                    op = False
                End If
                msg opnick$ & " deopped " & onick$ & "."
            Else
            j$ = j$ & m$
            End If
    Next i
End If
If cmd$ = "!+m " Then
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                If onick$ = un$ Then
                   muted = True
                End If
                msg opnick$ & " muted " & onick$ & "."
            Else
            j$ = j$ & m$
            End If
    Next i
End If
If cmd$ = "!-m " Then
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                If onick$ = un$ Then
                   muted = False
                End If
                msg opnick$ & " un-muted " & onick$ & "."
            Else
            j$ = j$ & m$
            End If
    Next i
End If
If cmd$ = "act " Then
msg rest$
                
End If
End Sub

Private Sub client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If Number = 11001 Then
Form1.Label3.Caption = ""
client.Close
End If
End Sub

Private Sub Form_Load()
loadprefs2
msg "Welcome to SqUaShComm v. " & App.Major & "." & App.Minor & App.Revision

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu mnufile

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If ext = False Then End

End Sub

Private Sub Form_Resize()
ResizeForm Me

End Sub

Private Sub List1_DblClick()
If Form2.List1.List(Form2.List1.ListIndex) = un$ Then Exit Sub
Dim frm As New Form4
Set frm = New Form4
frm.Caption = Form2.List1.List(Form2.List1.ListIndex)
frm.Label1.Caption = Form2.List1.List(Form2.List1.ListIndex)
frm.sock.LocalPort = lastprt + 1
frm.sock.Listen
send "con " & Form2.List1.List(Form2.List1.ListIndex) & "," & frm.sock.LocalIP & "," & frm.sock.LocalPort & "."
frm.Show
frm.RichTextBox2.SetFocus

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List1.ListIndex < 0 Then Exit Sub
If Button = 2 Then
If op = True Then
Form3.mnukick.Visible = True
Form3.mnuop.Visible = True
Form3.mnudeop.Visible = True
Form3.mnufilebar1.Visible = True
Form3.mnumute.Visible = True
Form3.mnuunmute.Visible = True
Else
Form3.mnukick.Visible = False
Form3.mnuop.Visible = False
Form3.mnudeop.Visible = False
Form3.mnufilebar1.Visible = False
Form3.mnumute.Visible = False
Form3.mnuunmute.Visible = False
End If
PopupMenu Form3.mnuuser
End If
End Sub



Private Sub mnuabout_Click()
Load frmAbout
frmAbout.Show

End Sub

Private Sub mnudisconnect_Click()
If server1 = True Then
rep% = MsgBox("Shut down server?", vbQuestion + vbYesNo, "Shut down server?")
If rep% = vbYes Then
End
Exit Sub
End If
Exit Sub
End If
On Error GoTo hell:
client.Close
DoEvents
ext = True
Unload Me
ext = False
Load Form1
Form1.Show
Unload Form4
Unload Form5
Unload Form6
Unload frmAbout
op = False
hell:
End Sub

Private Sub mnuedit_Click()
Load Form7
Form7.Show

End Sub

Private Sub mnuexit_Click()
End

End Sub

Private Sub mnuping_Click()

End Sub

Private Sub mnuhelp2_Click()
Load Form6
Form6.Show

End Sub

Private Sub mnuopt_Click()
Load Dialog
Dialog.Show 1, Form2

End Sub

Private Sub mnuupdate_Click()

End Sub

Private Sub ping_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo hell:
If NUM3 = 0 Then
NUM3 = 1
Load ping(NUM3)
ping(NUM3).Accept requestID
'//

Exit Sub
End If
For d = 1 To NUM3
    If ping(d).State <> 7 Then
        ping(d).Close
        ping(d).Accept requestID
'//

        Exit Sub
    End If
Next d
If NUM3 = musers Then Exit Sub
NUM3 = NUM3 + 1
Load ping(NUM3)
ping(NUM3).Accept requestID
'//

hell:
End Sub

Private Sub ping_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
ping(Index).GetData strdata
DoEvents
ping(Index).SendData strdata
DoEvents
End Sub

Private Sub ping2_Connect()
Timer1.Enabled = True
st = "[DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm][DanComm]"
ping2.SendData st
DoEvents

End Sub

Private Sub ping2_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
ping2.GetData strdata
DoEvents
Timer1.Enabled = False
ping2.Close
msg "User Response: " & tme & "ms", vbBlue
tme = 0
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i
If KeyCode = vbKeyReturn Then
If muted = True Then Exit Sub
        If RichTextBox1.Text = "" Then Exit Sub
                        msgs = msgs + 1
                If msgs > 2 Then
        RichTextBox1.Locked = True
            Timer3.Enabled = True
            msg "You cannot talk for 10 seconds (flood)."
            Timer2.Enabled = False
            Exit Sub
        End If
        If Left(RichTextBox1.Text, 6) = "/nick " Then
            If Len(RichTextBox1.Text) <= 6 Then Exit Sub
            On Error GoTo hell:
                send "!+n " & un$ & "," & Right(RichTextBox1.Text, Len(RichTextBox1.Text) - 6)
                RichTextBox1.Text = ""
                Exit Sub
        End If
        If Left(RichTextBox1.Text, 4) = "/me " Then
            If Len(RichTextBox1.Text) <= 6 Then Exit Sub
            On Error GoTo hell:
                send "act * " & un$ & Right(RichTextBox1.Text, Len(RichTextBox1.Text) - 3)
                RichTextBox1.Text = ""
                                
                Exit Sub
        End If

        send "msg " & un$ & ": " & RichTextBox1.Text
        RichTextBox1.Text = ""
End If
hell:

End Sub

Private Sub sock_Close(Index As Integer)
For i = 0 To List1.ListCount - 1
If user(Index).sn = List1.List(i) Then
List1.RemoveItem i
thesn$ = user(Index).sn
Exit For
End If
Next i
senduserlist
DoEvents
sendall "act * " & thesn$ & " left."
End Sub

Private Sub sock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
op = True
server1 = True
On Error GoTo hell:
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
For i = 0 To READINI("#block:")
    If READINI("block " & i & ":") = sock(Index).RemoteHostIP Then
        sock(Index).Close
        Exit Sub
    End If
Next i
THEFILE = n
If NUM = 0 Then
NUM = 1
Load sock(NUM)
sock(NUM).Accept requestID
'//

Exit Sub
End If
For d = 1 To NUM
    If sock(d).State <> 7 Then
        sock(d).Close
        sock(d).Accept requestID
'//

        Exit Sub
    End If
Next d
If NUM = musers Then Exit Sub
NUM = NUM + 1
Load sock(NUM)
sock(NUM).Accept requestID
'//

hell:
End Sub

Private Sub sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
sock(Index).GetData strdata
cmd$ = Left(strdata, 4)
rest$ = Right(strdata, Len(strdata) - 4)
If cmd$ = "msg " Then
sendall strdata
End If
If cmd$ = "pmg " Then
For i = 1 To Len(rest$)
    t$ = Mid(rest$, i, 1)
        If t$ = "," Then
                For d = 1 To NUM
                    If j$ = user(d).sn Then
                        send2 d, strdata
                        j$ = ""
                        Exit Sub
                    End If
                Next d
        Else
        j$ = j$ & t$
        End If
Next i
                
End If
If cmd$ = "con " Then
For i = 1 To Len(rest$)
    t$ = Mid(rest$, i, 1)
        If t$ = "," Then
                For d = 1 To NUM
                    If j$ = user(d).sn Then
                        send2 d, strdata
                        j$ = ""
                        Exit Sub
                    End If
                Next d
        Else
        j$ = j$ & t$
        End If
Next i
                
End If
If cmd$ = "png " Then
    For i = 1 To NUM
        If user(i).sn = rest$ Then
            send2 Index, "png " & sock(i).RemoteHostIP
        End If
    Next i
End If
If cmd$ = "add " Then
                    If validatenick(rest$) = False Then
                    send2 Index, "!-n Nick not accepted."
                    sock(Index).Close
                    Exit Sub
                    End If
    For i = 0 To List1.ListCount - 1
        If LCase(rest$) = LCase(List1.List(i)) Then
            send2 Index, "dis User name already in use."
            sock(Index).Close
            Exit Sub
        End If
    Next i
List1.AddItem rest$
user(Index).sn = rest$
user(Index).chanlev = 0
If rest$ = un$ Then user(Index).chanlev = 5
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
For i = 0 To READINI("#access:")
    If LCase(READINI("access " & i & ":")) = LCase(rest$) Then
        If user(Index).chanlev = 0 Then sendall "!+o " & rest$ & "," & un$
        user(Index).chanlev = 5
        Exit For
    End If
Next i
THEFILE = n
senduserlist
DoEvents
sendall "act * " & rest$ & " joined."
  End If
  '// new room commands
  If cmd$ = "!+o " Then
  For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                For d = 1 To NUM
                    If onick$ = user(d).sn Then
                    If onick$ = un$ And server1 = True Then Exit Sub
                        If user(d).chanlev = 5 Then Exit Sub
                        user(d).chanlev = 5
                    End If
                Next d
            Else
            j$ = j$ & m$
            End If
    Next i
  sendall strdata
End If
If cmd$ = "!-o " Then
  For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                onick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                For d = 1 To NUM
                    If onick$ = user(d).sn Then
                    If onick$ = un$ And server1 = True Then Exit Sub
                    If user(d).chanlev = 0 Then Exit Sub
                        user(d).chanlev = 0
                    End If
                Next d
            Else
            j$ = j$ & m$
            End If
    Next i
  sendall strdata
End If
If cmd$ = "!+k " Then
  For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                knick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                    For d = 1 To NUM
                    If knick$ = user(d).sn Then
                    If knick$ = un$ And server1 = True Then Exit Sub
                        send2 d, "dis You were kicked by " & opnick$ & "."
                        sock(d).Close
                    End If
                    Next d
                    For d = 0 To List1.ListCount - 1
                    If knick$ = List1.List(d) Then
                        List1.RemoveItem d
                        Exit For
                    End If
                    Next d
                msg opnick$ & " kicked " & knick$ & "."

            Else
            j$ = j$ & m$
            End If
    Next i
  sendall strdata
  senduserlist
End If
If cmd$ = "!+m " Then
  For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                mnick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                    For d = 1 To NUM
                    If mnick$ = user(d).sn Then
                    If mnick$ = un$ And server1 = True Then Exit Sub
                    If user(d).chanlev = -1 Then Exit Sub
                    user(d).chanlev = -1
                    End If
                    Next d

            Else
            j$ = j$ & m$
            End If
    Next i
  sendall strdata
End If
If cmd$ = "!-m " Then
  For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
                If m$ = "," Then
                mnick$ = j$
                opnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                    For d = 1 To NUM
                    If mnick$ = user(d).sn Then
                    If mnick$ = un$ And server1 = True Then Exit Sub
                    If user(d).chanlev = 0 Then Exit Sub
                    user(d).chanlev = 0
                    End If
                    Next d

            Else
            j$ = j$ & m$
            End If
    Next i
  sendall strdata
End If
'// new nickname...
  If cmd$ = "!+n " Then
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
            If m$ = "," Then
                oldnick$ = j$
                newnick$ = Right(rest$, Len(rest$) - Len(j$) - 1)
                    For d = 0 To List1.ListCount - 1
                        If LCase(newnick$) = LCase(List1.List(d)) Then
                            send2 Index, "!-n User name already in use."
                            Exit Sub
                        ElseIf oldnick$ = List1.List(d) Then
                                    
                            oldind = d
                        End If
                        
                    Next d
                    If validatenick(newnick$) = False Then
                    send2 Index, "!-n Nick not accepted."
                    Exit Sub
                    Else
                    send2 Index, "!/n " & newnick$
                    List1.RemoveItem oldind
                List1.AddItem newnick$
                user(Index).sn = newnick$
                DoEvents
                senduserlist
                Exit Sub
                End If
        Else
        j$ = j$ & m$
        End If
                
    Next i
End If
If cmd$ = "act " Then
sendall strdata
                
End If
End Sub

Private Sub Timer1_Timer()
tme = tme + 1

End Sub

Private Sub Timer2_Timer()
msgs = 0
End Sub

Private Sub Timer3_Timer()
tim = tim + 1
RichTextBox1.Locked = True
If tim = 10 Then
Timer3.Enabled = False
Timer2.Enabled = True
msg "You can talk again now."
    RichTextBox1.Locked = False
    tim = 0
    msgs = 0
End If
End Sub
