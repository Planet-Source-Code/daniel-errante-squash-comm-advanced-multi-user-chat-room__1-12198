VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SqUaShComm Options"
   ClientHeight    =   4860
   ClientLeft      =   2760
   ClientTop       =   3780
   ClientWidth     =   4710
   Icon            =   "dcommopt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   4710
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "dcommopt.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Check1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Ignore List"
      TabPicture(1)   =   "dcommopt.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Access List"
      TabPicture(2)   =   "dcommopt.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Block List"
      TabPicture(3)   =   "dcommopt.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Filtering"
      TabPicture(4)   =   "dcommopt.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label7"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Check2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Text6"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   39
         Text            =   "*"
         Top             =   840
         Width           =   495
      End
      Begin VB.Frame Frame4 
         Caption         =   "Bad Word List:"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   33
         Top             =   1200
         Width           =   4215
         Begin VB.ListBox List4 
            Height          =   1815
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   37
            Top             =   240
            Width           =   3975
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   800
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1680
            TabIndex        =   35
            Top             =   2280
            Width           =   800
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3240
            TabIndex        =   34
            Top             =   2280
            Width           =   800
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Filter Bad Words"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   3375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Access List (server only)"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   4215
         Begin VB.CommandButton Command11 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3240
            TabIndex        =   31
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1680
            TabIndex        =   30
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   3000
            Width           =   800
         End
         Begin VB.ListBox List3 
            Height          =   2400
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   28
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ignore List"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   4215
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1680
            TabIndex        =   25
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   3000
            Width           =   800
         End
         Begin VB.ListBox List1 
            Height          =   2400
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   23
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Block List (server only)"
         Height          =   3495
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   4215
         Begin VB.CommandButton Command4 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3240
            TabIndex        =   21
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1680
            TabIndex        =   20
            Top             =   3000
            Width           =   800
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Add"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   3000
            Width           =   800
         End
         Begin VB.ListBox List2 
            Height          =   2400
            Left            =   120
            MultiSelect     =   2  'Extended
            TabIndex        =   18
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Text            =   "Welcome to SqUaShComm"
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Font"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Font"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   3120
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable Private Messages"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Bad Word Char:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   38
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   $"dcommopt.frx":0396
         Height          =   915
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Chat Background:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Chat Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Private Message Background:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Private Message To Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Private Message From Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   2415
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim rep$
rep$ = InputBox$("Add user name to ignore list:", "Add User:")
If rep$ = "" Then Exit Sub
List1.AddItem rep$

End Sub

Private Sub Command10_Click()
Dim d As Integer
d = List3.ListCount - 1
Do Until d = -1
If List3.Selected(d) = True Then List3.RemoveItem d
d = d - 1
Loop
End Sub

Private Sub Command11_Click()
Dim rep%
rep% = MsgBox("Do you want to clear access list?", vbQuestion + vbYesNo, "Clear access list?")
If rep% = vbYes Then List3.Clear

End Sub

Private Sub Command12_Click()
Dim rep%
rep% = MsgBox("Do you want to clear Bad Word list?", vbQuestion + vbYesNo, "Clear Bad Word list?")
If rep% = vbYes Then List4.Clear

End Sub

Private Sub Command13_Click()
Dim d As Integer
d = List4.ListCount - 1
Do Until d = -1
If List4.Selected(d) = True Then List4.RemoveItem d
d = d - 1
Loop
End Sub

Private Sub Command14_Click()
Dim rep$
rep$ = InputBox$("Add a word to Bad Word List:", "Add Bad Word:")
If rep$ = "" Then Exit Sub
List4.AddItem rep$

End Sub

Private Sub Command2_Click()
Dim d As Integer
d = List1.ListCount - 1
Do Until d = -1
If List1.Selected(d) = True Then List1.RemoveItem d
d = d - 1
Loop
End Sub

Private Sub Command3_Click()
Dim rep%
rep% = MsgBox("Do you want to clear ignore list?", vbQuestion + vbYesNo, "Clear ignore list?")
If rep% = vbYes Then List1.Clear

End Sub

Private Sub Command4_Click()
Dim rep%
rep% = MsgBox("Do you want to clear block list?", vbQuestion + vbYesNo, "Clear block list?")
If rep% = vbYes Then List2.Clear

End Sub

Private Sub Command5_Click()
Dim d As Integer
d = List2.ListCount - 1
Do Until d = -1
If List2.Selected(d) = True Then List2.RemoveItem d
d = d - 1
Loop
End Sub

Private Sub Command6_Click()
Dim rep$
rep$ = InputBox$("Add an IP address to block list:", "Add IP:")
If rep$ = "" Then Exit Sub
List2.AddItem rep$

End Sub

Private Sub Command7_Click()
On Error GoTo hell:
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
Dialog.Text2.Font = CommonDialog1.FontName
Dialog.Text2.FontSize = CommonDialog1.FontSize
Dialog.Text2.FontBold = CommonDialog1.FontBold
Dialog.Text2.FontItalic = CommonDialog1.FontItalic

hell:

End Sub

Private Sub Command8_Click()
On Error GoTo hell:
CommonDialog1.Flags = cdlCFBoth
CommonDialog1.ShowFont
Dialog.Text4.Font = CommonDialog1.FontName
Dialog.Text4.FontSize = CommonDialog1.FontSize
Dialog.Text4.FontBold = CommonDialog1.FontBold
Dialog.Text4.FontItalic = CommonDialog1.FontItalic
Dialog.Text5.Font = CommonDialog1.FontName
Dialog.Text5.FontSize = CommonDialog1.FontSize
Dialog.Text5.FontBold = CommonDialog1.FontBold
Dialog.Text5.FontItalic = CommonDialog1.FontItalic

hell:

End Sub

Private Sub Command9_Click()
Dim rep$
rep$ = InputBox$("Add a user name to Access List:", "Add IP:")
If rep$ = "" Then Exit Sub
List3.AddItem rep$

End Sub

Private Sub Form_Load()
loadprefs
Text4.Text = un$ & ": hi"
Text5.Text = "user: hello"

End Sub

Private Sub OKButton_Click()
saveprefs
DoEvents
loadprefs2
DoEvents
Unload Me

End Sub

Private Sub Text1_DblClick()
On Error GoTo hell:
Text2.SelLength = 0
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Text1.BackColor = CommonDialog1.Color
Text2.BackColor = CommonDialog1.Color
List1.BackColor = CommonDialog1.Color
List2.BackColor = CommonDialog1.Color
List3.BackColor = CommonDialog1.Color
List4.BackColor = CommonDialog1.Color
Text6.BackColor = CommonDialog1.Color

hell:
End Sub

Private Sub Text2_DblClick()
On Error GoTo hell:
Text2.SelLength = 0
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Text2.ForeColor = CommonDialog1.Color
List1.ForeColor = CommonDialog1.Color
List2.ForeColor = CommonDialog1.Color
List3.ForeColor = CommonDialog1.Color
List4.ForeColor = CommonDialog1.Color
Text6.ForeColor = CommonDialog1.Color

hell:
End Sub

Private Sub Text3_DblClick()
On Error GoTo hell:
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Text3.BackColor = CommonDialog1.Color
Text4.BackColor = CommonDialog1.Color
Text5.BackColor = CommonDialog1.Color

hell:
End Sub

Private Sub Text4_DblClick()
On Error GoTo hell:
Text4.SelLength = 0
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Text4.ForeColor = CommonDialog1.Color

hell:
End Sub

Private Sub Text5_DblClick()
On Error GoTo hell:
Text5.SelLength = 0
CommonDialog1.CancelError = True
CommonDialog1.ShowColor
Text5.ForeColor = CommonDialog1.Color

hell:
End Sub
