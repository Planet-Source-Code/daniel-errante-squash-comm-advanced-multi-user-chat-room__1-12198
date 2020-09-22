Attribute VB_Name = "Module1"
Public un$
Public server$
Public host$
Public NUM As Integer
Public NUM2 As Integer
Public NUM3 As Integer
Const maxusers = 100
Public musers As String
Public ignore As String
Public lastprt As Integer
Public txtclr As Long
Public txtfnt As String
Public txtbld As Boolean
Public txtsze As Integer
Public txtit As Boolean
'//
Public pmclr As Long
Public pmclr2 As Long
Public pmfnt As String
Public pmbld As Boolean
Public pmsze As Integer
Public pmit As Boolean
Public pms As String
Public op As Boolean
Public server1 As Boolean
Public muted As Boolean
Public ext As Boolean
Public msgs As Integer
Public tim As Integer

Type clientuser

    sn As String
    chanlev As Integer
    
End Type

Public user(maxusers) As clientuser

Function send(dat As String)
If Form2.client.State = 7 Then Form2.client.SendData dat
DoEvents

End Function

Function sendhost(dat As String)
If Form4.sock.State = 7 Then Form4.sock.SendData dat
DoEvents

End Function

Function sendclient(dat As String)
If Form5.client.State = 7 Then Form5.client.SendData dat
DoEvents

End Function

Function msg(message As String, Optional clr As String = vbBlack, Optional bld As Boolean = False, Optional sze As Long = 10)

Form2.RichTextBox2.SelStart = Len(Form2.RichTextBox2.Text)
Form2.RichTextBox2.SelColor = txtclr
Form2.RichTextBox2.SelFontSize = txtsze
Form2.RichTextBox2.SelBold = txtbld
Form2.RichTextBox2.SelItalic = txtit
Form2.RichTextBox2.SelFontName = txtfnt
Form2.RichTextBox2.SelText = message & vbCrLf
Form2.RichTextBox2.SelStart = Len(Form2.RichTextBox2.Text)
End Function


Function send2(prt2, dat As String)
If Form2.sock(prt2).State = 7 Then Form2.sock(prt2).SendData dat
DoEvents

End Function

Function senduserlist()
    For d = 0 To Form2.List1.ListCount - 1
        userlist$ = userlist$ & Form2.List1.List(d) & ","
    Next d
For i = 1 To NUM
If Form2.sock(i).State = 7 Then send2 i, "lst " & userlist$
Next i
End Function

Function sendall(dat As String)
For i = 1 To NUM
If Form2.sock(i).State = 7 Then send2 i, dat
Next i
End Function


Function saveprefs()
On Error GoTo hell:
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
CREATEINI
SAVEINI "bgcolor:", Dialog.Text1.BackColor
SAVEINI "textcolor:", Dialog.Text2.ForeColor
SAVEINI "pms:", Dialog.Check1.Value
SAVEINI "pmbg:", Dialog.Text3.BackColor
SAVEINI "pmto:", Dialog.Text4.ForeColor
SAVEINI "pmfrom:", Dialog.Text5.ForeColor
SAVEINI "textbld:", Dialog.Text2.FontBold
SAVEINI "textitalic:", Dialog.Text2.FontItalic
SAVEINI "textfnt:", Dialog.Text2.FontName
SAVEINI "textsize:", Dialog.Text2.FontSize
'//
SAVEINI "pmbld:", Dialog.Text4.FontBold
SAVEINI "pmitalic:", Dialog.Text4.FontItalic
SAVEINI "pmfnt:", Dialog.Text4.FontName
SAVEINI "pmsize:", Dialog.Text4.FontSize
'//
SAVEINI "#ignore:", Dialog.List1.ListCount - 1
    For i = 0 To Dialog.List1.ListCount - 1
        SAVEINI "ignore " & i & ":", Dialog.List1.List(i)
    Next i
SAVEINI "#block:", Dialog.List2.ListCount - 1
        For i = 0 To Dialog.List2.ListCount - 1
        SAVEINI "block " & i & ":", Dialog.List2.List(i)
        Next i
        SAVEINI "#access:", Dialog.List3.ListCount - 1
        For i = 0 To Dialog.List3.ListCount - 1
        SAVEINI "access " & i & ":", Dialog.List3.List(i)
        Next i
THEFILE = App.Path & "\badwords.ini"
CREATEINI
SAVEINI "FILTER:", Dialog.Check2.Value
SAVEINI "BW CHAR:", Dialog.Text6.Text
SAVEINI "#BW:", Dialog.List4.ListCount - 1
For i = 0 To Dialog.List4.ListCount - 1
    SAVEINI "BW " & i & ":", Dialog.List4.List(i)
Next i
THEFILE = n
hell:

End Function

Function loadprefs()
On Error GoTo hell:
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
Dialog.Text1.BackColor = READINI("bgcolor:")
Dialog.Text2.BackColor = READINI("bgcolor:")
Dialog.List1.BackColor = READINI("bgcolor:")
Dialog.List2.BackColor = READINI("bgcolor:")
Dialog.List3.BackColor = READINI("bgcolor:")
Dialog.Text2.ForeColor = READINI("textcolor:")
Dialog.Text2.FontName = READINI("textfnt:")
Dialog.Text2.FontSize = READINI("textsize:")
Dialog.Text2.FontBold = READINI("textbld:")
Dialog.Text2.FontItalic = READINI("textitalic:")
Dialog.Text3.ForeColor = READINI("textcolor:")
Dialog.List1.ForeColor = READINI("textcolor:")
Dialog.List2.ForeColor = READINI("textcolor:")
Dialog.List3.ForeColor = READINI("textcolor:")
Dialog.Text3.BackColor = READINI("pmbg:")
Dialog.Text4.BackColor = READINI("pmbg:")
Dialog.Text5.BackColor = READINI("pmbg:")
Dialog.Text4.ForeColor = READINI("pmto:")
Dialog.Text5.ForeColor = READINI("pmfrom:")
Dialog.Text5.FontName = READINI("pmfnt:")
Dialog.Text5.FontSize = READINI("pmsize:")
Dialog.Text5.FontBold = READINI("pmbld:")
Dialog.Text5.FontItalic = READINI("pmitalic:")
Dialog.Text4.FontName = READINI("pmfnt:")
Dialog.Text4.FontSize = READINI("pmsize:")
Dialog.Text4.FontBold = READINI("pmbld:")
Dialog.Text4.FontItalic = READINI("pmitalic:")
Dialog.Text6.BackColor = READINI("bgcolor:")
Dialog.Text6.ForeColor = READINI("textcolor:")
Dialog.List4.BackColor = READINI("bgcolor:")
Dialog.List4.ForeColor = READINI("textcolor:")
Dialog.List1.Clear
Dialog.Check1.Value = READINI("pms:")
For i = 0 To READINI("#ignore:")
    Dialog.List1.AddItem READINI("ignore " & i & ":")
Next i
Dialog.List2.Clear
For i = 0 To READINI("#block:")
    Dialog.List3.AddItem READINI("block " & i & ":")
Next i
Dialog.List3.Clear
For i = 0 To READINI("#access:")
    Dialog.List3.AddItem READINI("access " & i & ":")
Next i
THEFILE = App.Path & "\badwords.ini"
   Dialog.Check2.Value = READINI("FILTER:")
   Dialog.Text6.Text = READINI("BW CHAR:")
   For i = 0 To READINI("#BW:")
    Dialog.List4.AddItem READINI("BW " & i & ":")
    Next i
THEFILE = n
hell:

End Function

Function loadprefs2()
On Error GoTo hell:
n = THEFILE
THEFILE = App.Path & "\dancomm.ini"
Form2.RichTextBox1.SelStart = 0
Form2.RichTextBox1.SelLength = Len(Form2.RichTextBox1.Text)
Form2.RichTextBox1.BackColor = READINI("bgcolor:")
Form2.RichTextBox1.SelStart = Len(Form2.RichTextBox1.Text)
Form2.RichTextBox2.SelStart = 0
Form2.RichTextBox2.SelLength = Len(Form2.RichTextBox2.Text)
Form2.RichTextBox2.BackColor = READINI("bgcolor:")
Form2.RichTextBox2.SelStart = Len(Form2.RichTextBox2.Text)
Form2.List1.BackColor = READINI("bgcolor:")
Form2.List1.ForeColor = READINI("textcolor:")
'//
txtclr = READINI("textcolor:")
txtbld = READINI("textbld:")
txtit = READINI("textitalic:")
txtfnt = READINI("textfnt:")
txtsze = READINI("textsize:")
'//
Form2.RichTextBox1.SelColor = txtclr
Form2.RichTextBox1.SelFontSize = txtsize
Form2.RichTextBox1.SelBold = txtbld
Form2.RichTextBox1.SelFontName = txtfnt
Form2.RichTextBox1.SelItalic = txtit
'//
pms = READINI("pms:")
pmclr = READINI("pmto:")
pmclr2 = READINI("pmfrom:")
pmbld = READINI("pmbld:")
pmit = READINI("pmitalic:")
pmfnt = READINI("pmfnt:")
pmsze = READINI("pmsize:")
THEFILE = n
hell:

End Function


Function validatenick(thenick As String) As Boolean
If Len(thenick) > 30 Then
validatenick = False
Exit Function
End If
Open App.Path & "\badnicks.ini" For Input As #7
Do Until EOF(7)
    Line Input #7, lineoftext$
    alltext$ = alltext$ & lineoftext$
Loop
Close #7
nontxt = alltext$
For i = 1 To Len(thenick)
m$ = Mid(thenick, i, 1)
    For d = 1 To Len(nontxt)
        t$ = Mid(nontxt, d, 1)
            If t$ = m$ Then
            validatenick = False
            Exit Function
            End If
    Next d

Next i
THEFILE = App.Path & "\badwords.ini"
If READINI("FILTER:") = "0" Then GoTo nd:
ch$ = READINI("BW CHAR:")
For i = 0 To READINI("#BW:")
t$ = READINI("BW " & i & ":")
r = Len(t$)
            For d = 1 To Len(thenick)
            m$ = Mid(thenick, d, r)
    If LCase(t$) = LCase(m$) And t$ <> "" And m$ <> "" Then
        validatenick = False
        THEFILE = n
        Exit Function
    End If
        Next d
Next i
nd:
THEFILE = n
validatenick = True
End Function

Function validatemsg(themsg As String) As String
n = THEFILE
THEFILE = App.Path & "\badwords.ini"
If READINI("FILTER:") = "0" Then GoTo nd:
ch$ = READINI("BW CHAR:")
For i = 0 To READINI("#BW:")
t$ = READINI("BW " & i & ":")
r = Len(t$)
            For d = 1 To Len(themsg)
            m$ = Mid(themsg, d, r)
    If LCase(t$) = LCase(m$) Then
        themsg = Left(themsg, d - 1) & String(r, ch$) & Right(themsg, Len(themsg) - d - r + 1)
    End If
        Next d
Next i
nd:
THEFILE = n
validatemsg = themsg
End Function
