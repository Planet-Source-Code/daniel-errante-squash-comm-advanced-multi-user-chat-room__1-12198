Attribute VB_Name = "readinimod"
'//READINI
'//BY DANIEL ERRANTE
'//danoph@hotmail.com
'//Copyright c 1999 Daniel Errante. All
'     rights reserved.
'//www.nowresources.com/danoph
'//Find other cool apps above
'//
'//
'//DIRECTIONS FOR READINI
'//
'//
'//READS INI OR INF FILES YOU CREATE USI
'     NG THIS MODULE
'//To use, declare the following in your


'     form where you call the function READINI
    '
    '//This allows for multiple INI file rea
    '     ding
    '//For Example:
    '//THEFILE = App.Path & "\options.inf"
    '//Then, call the function. The function
    '     uses the variable you declared so
    '//you have to declare it before you use
    '     the function
    '//To call the function, use this syntax
    '     :
    '//variable = READINI(COMMAND As String)
    '
    '//Example:
    '//Text1.text = READINI("NAME: "")
    '//If the text of THEFILE equaled "NAME:
    '     Daniel Errante (danoph@hotmail.com)"
    '//Then Text1.Text would equal "Daniel E
    '     rrante (danoph@hotmail.com)"
    Public THEFILE As String


Public Function CREATEINI()
    Open THEFILE For Output As #1
    Print #1, "//" & THEFILE
    Close #1
End Function


Public Function READINI(COMMAND As String) As String
    If COMMAND = "" Then READINI = Null
    COMMANDLENGTH = Len(COMMAND)
Open_file:
    Close #1 '//just To make sure
    On Error GoTo err: '//file Not found?
    Open THEFILE For Input As #1 'opens previously declared file


    Do Until EOF(1) 'search For command until End of file
        Line Input #1, lineoftext$ 'gets a new line of the file


        If lineoftext$ <> "" Then
            cmnd$ = Left(lineoftext$, COMMANDLENGTH) '//gets command


            If cmnd$ = COMMAND Then '//if we got a winner:
                VALU$ = Right(lineoftext$, Len(lineoftext$) - COMMANDLENGTH) '//this is the value of the command
                READINI = VALU$ '//return the value of the command
            End If
        End If
    Loop '//do until command is found
    Close #1
    Exit Function
err:


    If err.Number = 53 Then
        CREATEINI
        GoTo Open_file
    End If
End Function


Public Function SAVEINI(COMMAND As String, VALU As String)
    '//Returns true if save successful
    On Error GoTo err:
    filetext$ = PRE_OPEN
    On Error GoTo err:
    Close #2
    Open THEFILE For Output As #2
    Print #2, filetext$ & COMMAND & VALU
    Close #2
    Exit Function
err:
End Function


Public Function PRE_OPEN() As String
    '//Opens file and returns the text
    Close #3
    On Error GoTo err:
    Open THEFILE For Input As #3


    Do Until EOF(3)
        Line Input #3, lineoftext$
        alltext$ = alltext$ & lineoftext$ & vbCrLf
    Loop
    Close #3
    PRE_OPEN = alltext$
    Exit Function
err:


    If err.Number = 53 Then
        PRE_OPEN = "file Not found"
    End If
End Function


Public Function FIND(txt As String, txtcmd As String, numcmd As String) As Boolean
'//format "USER1:DANIEL"
For i = 1 To READINI(numcmd)
DoEvents
If READINI(txtcmd & i & ":") = txt Then
FIND = True
Exit Function
End If
DoEvents
Next i
End Function
