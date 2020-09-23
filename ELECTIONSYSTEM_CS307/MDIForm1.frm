VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10125
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "MDIForm1.frx":0000
   MousePointer    =   99  'Custom
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5520
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2436
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":426C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":46BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6398
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":67EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":84C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":90B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9992
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A66C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B716
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C568
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":D612
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7395
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3675
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   0
         Picture         =   "MDIForm1.frx":DA64
         ScaleHeight     =   1215
         ScaleWidth      =   3615
         TabIndex        =   2
         Top             =   0
         Width           =   3615
      End
      Begin MSComctlLib.ListView lshortcut 
         Height          =   8055
         Left            =   0
         TabIndex        =   1
         Top             =   1560
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   14208
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         _Version        =   393217
         ColHdrIcons     =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student Info"
            Object.Width           =   2540
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Election Result"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Election Winners"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Add new Candidate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Log - out Admin"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -120
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblyear 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Menu myright 
      Caption         =   "my"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuclose2 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu myright1 
      Caption         =   "my2"
      Visible         =   0   'False
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close Active Form"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdelecinfo_Click()
On Error Resume Next
Unload MDIForm1.ActiveForm
formstudent.Show
End Sub

Private Sub cmdelecproper_Click()

End Sub

Private Sub cmdelecres_Click()
Form4.Show
End Sub

Private Sub cmdnewofficer_Click()

End Sub

Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub

Private Sub lshortcut_DblClick()
Dim askmelogout As Single
On Error Resume Next
Unload MDIForm1.ActiveForm
Select Case Me.lshortcut.SelectedItem.Key
    Case "addnew": formapply.Show
    Case "elecres": Form4.Show
    Case "electwin": formwinner.Show: Unload Me.ActiveForm: Me.Hide
    Case "find": formfind.Show
    Case "logout":
        askmelogout = MsgBox("Continue log - out?", vbOKCancel + vbQuestion, "SSITE")
            If askmelogout = vbOK Then
                Form1.Show: Unload Me.ActiveForm: Me.Hide
            End If
    Case "studinfo": formstudent.Show
    Case "report": formreport.Show
    Case "Timer":
        With formtimer
            .adotime.Recordset.Filter = "DateVote LIKE '" & Format$(Date, "MM/DD/YYYY") & "'"
            If .adotime.Recordset.RecordCount > 0 Then
            .Show
            Else
                MsgBox "No registered voter by this date....", vbOKOnly + vbExclamation, "SSITE"
                Exit Sub
            End If
        End With
    Case "candidate": Unload MDIForm1.ActiveForm: formcandidate.Show
    Case "change": Unload MDIForm1.ActiveForm: formchange.Show
    Case "custom": Unload MDIForm1.ActiveControl: formyear.Show
End Select
End Sub

Private Sub lshortcut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu myright
End If
End Sub

Private Sub MDIForm_Load()
    With Me.lshortcut
        Set .SmallIcons = ImageList2
        Set .Icons = ImageList2
        'For Sales
        .ListItems.Add , "addnew", "Add new Candidate", 1, 1
        .ListItems.Add , "elecres", "Election Result", 2, 2
        .ListItems.Add , "electwin", "Election Winner", 6, 6
        .ListItems.Add , "studinfo", "Student Info", 7, 7
        .ListItems.Add , "find", "Find", 3, 3
        .ListItems.Add , "logout", "Log-out Admin", 8, 8
        .ListItems.Add , "report", "Report", 11, 11
        .ListItems.Add , "Timer", "Voting Status", 12, 12
        .ListItems.Add , "candidate", "View/Delete Candidates", 13, 13
        .ListItems.Add , "change", "Change My Password", 14, 14
        .ListItems.Add , "custom", "Customize", 15, 15
    End With
    Me.lblyear.Caption = Format$(Date, "YYYY")
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu myright1
End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim askmequit As Single
    askmequit = MsgBox("Close application?", vbYesNo + vbQuestion, "SSITE")
        If askmequit = vbYes Then
            End
        Else
            Cancel = 1
        End If
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    Me.WindowState = 2
End Sub

Private Sub mnuclose2_Click()
On Error Resume Next
Unload MDIForm1.ActiveForm
End Sub
Private Sub mnuclose_Click()
On Error Resume Next
Unload MDIForm1.ActiveForm
End Sub

Private Sub mnuexit_Click()
Dim askmequit As Single
    askmequit = MsgBox("Close application?", vbYesNo + vbQuestion, "SSITE")
        If askmequit = vbYes Then
            End
        Else
            Cancel = 1
        End If
End Sub

Private Sub mnulogout_Click()
Dim askmelogout As Single
askmelogout = MsgBox("Continue log - out?", vbOKCancel + vbQuestion, "SSITE")
            If askmelogout = vbOK Then
                Form1.Show: Unload Me.ActiveForm: Me.Hide
            End If
End Sub

Private Sub mnuOpen_Click()
lshortcut_DblClick
End Sub


Private Sub mnurefresh_Click()
On Error Resume Next
MDIForm1.ActiveForm.Refresh
End Sub
