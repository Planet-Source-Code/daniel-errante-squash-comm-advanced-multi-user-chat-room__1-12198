VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuuser 
      Caption         =   "&User"
      Begin VB.Menu mnuping 
         Caption         =   "&Ping"
      End
      Begin VB.Menu mnupm 
         Caption         =   "&Private Message"
      End
      Begin VB.Menu mnufilebar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnumute 
         Caption         =   "&Mute"
      End
      Begin VB.Menu mnuunmute 
         Caption         =   "&Un-mute"
      End
      Begin VB.Menu mnuop 
         Caption         =   "&Op"
      End
      Begin VB.Menu mnudeop 
         Caption         =   "&Deop"
      End
      Begin VB.Menu mnukick 
         Caption         =   "&Kick"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuprivatemsg_Click()

End Sub

Private Sub mnuban_Click()
send "!+b " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$

End Sub

Private Sub mnudeop_Click()
send "!-o " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$

End Sub

Private Sub mnukick_Click()
send "!+k " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$

End Sub

Private Sub mnumute_Click()
send "!+m " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$

End Sub

Private Sub mnuop_Click()
send "!+o " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$

End Sub

Private Sub mnuping_Click()
send "png " & Form2.List1.List(Form2.List1.ListIndex)

End Sub

Private Sub mnupm_Click()
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

Private Sub mnuunmute_Click()
send "!-m " & Form2.List1.List(Form2.List1.ListIndex) & "," & un$
End Sub
