VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formcandidate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Candidates"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc adodeleting 
      Height          =   330
      Left            =   120
      Top             =   9120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"formcan.frx":0000
      OLEDBString     =   $"formcan.frx":009A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM candidatesinfo"
      Caption         =   "adocan"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtdele 
      DataField       =   "StudentNo"
      DataSource      =   "adodeleting"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   8760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmddelcan 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Picture         =   "formcan.frx":0134
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "formcan.frx":0576
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox txtstud 
      DataField       =   "StudentNo"
      DataSource      =   "adocandidate"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adocandidate 
      Height          =   375
      Left            =   9720
      Top             =   8280
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"formcan.frx":1240
      OLEDBString     =   $"formcan.frx":12DA
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM query_Candidate"
      Caption         =   "adocandidate"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   -240
      Picture         =   "formcan.frx":1374
      ScaleHeight     =   1500
      ScaleWidth      =   9000
      TabIndex        =   6
      Top             =   0
      Width           =   9060
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   8655
      Begin VB.Label l4 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label l3 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Course/Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label l2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label l1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Student No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "formcandidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itm
Public Sub candidates()
Me.lv1.ListItems.Clear
With Me.adocandidate.Recordset
    Me.adocandidate.Refresh
    .Filter = "SchoolYear Like '" & MDIForm1.lblyear.Caption & "'"
        If .RecordCount <> 0 Then
        .MoveFirst
    Do Until .EOF
      Set itm = Me.lv1.ListItems.Add(, , .Fields!StudentNo)
          itm.SubItems(1) = .Fields("Surname") & ", " & .Fields("firstname")
          itm.SubItems(2) = .Fields("Course") & " / " & .Fields("Year")
          itm.SubItems(3) = .Fields("Title")
        .MoveNext
    Loop
        End If
    End With
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmddelcan_Click()
Dim askmeqng As Single
    If Me.lv1.SelectedItem.Text = "" Then
        MsgBox "Please select for deletion", vbOKOnly + vbExclamation, "SSITE"
        Exit Sub
    End If
    askmeqng = MsgBox("Are you sure you want to delete this candidate?", vbYesNo + vbQuestion, "SSITE")
        If askmeqng = vbYes Then
            With Me.adodeleting.Recordset
                .Delete
                .MoveNext
                If .EOF Then Me.adodeleting.Refresh
            End With
        End If
    Me.adocandidate.Refresh
    Me.adodeleting.Refresh
    Call candidates
End Sub

Private Sub Form_Activate()
Me.Move 0, 0
Call candidates
End Sub

Private Sub Form_Load()
With Me.lv1
    .ColumnHeaders(1).Width = l1.Width
    .ColumnHeaders(2).Width = l2.Width
    .ColumnHeaders(3).Width = l3.Width
    .ColumnHeaders(4).Width = l4.Width
End With
End Sub

Private Sub lv1_Click()
With Me.adodeleting.Recordset
        .Filter = "StudentNo LIKE '" & Me.lv1.SelectedItem.Text & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
End With
With Me.adocandidate.Recordset
        .Filter = "StudentNo LIKE '" & Me.lv1.SelectedItem.Text & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
End With
End Sub
