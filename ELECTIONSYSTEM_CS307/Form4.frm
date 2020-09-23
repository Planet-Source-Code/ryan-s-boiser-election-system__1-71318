VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "SSITE "
   ClientHeight    =   11490
   ClientLeft      =   -7140
   ClientTop       =   -2235
   ClientWidth     =   16545
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   MouseIcon       =   "Form4.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11490
   ScaleWidth      =   16545
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Form4.frx":030A
      ScaleHeight     =   825
      ScaleWidth      =   12225
      TabIndex        =   21
      Top             =   0
      Width           =   12255
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9600
      Width           =   2535
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6000
      Picture         =   "Form4.frx":B006
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9600
      Width           =   2535
   End
   Begin VB.TextBox txtres 
      DataField       =   "VoteCount"
      DataSource      =   "adowinner"
      Height          =   285
      Left            =   18720
      TabIndex        =   18
      Top             =   10800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView lpres 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSAdodcLib.Adodc adowinner 
      Height          =   375
      Left            =   18120
      Top             =   10680
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   $"Form4.frx":BCD0
      OLEDBString     =   $"Form4.frx":BD6A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM res_query"
      Caption         =   ""
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
   Begin MSComctlLib.ListView lvf 
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvia 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvea 
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsg 
      Height          =   1575
      Left            =   6000
      TabIndex        =   13
      Top             =   1200
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lsp 
      Height          =   1575
      Left            =   6000
      TabIndex        =   14
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ltres 
      Height          =   1575
      Left            =   6000
      TabIndex        =   15
      Top             =   5040
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lau 
      Height          =   1575
      Left            =   6000
      TabIndex        =   16
      Top             =   6960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firsname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lbm 
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   8880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Firstname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Vote Counts"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label l9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Board Members"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Width           =   5385
   End
   Begin VB.Label l8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auditor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   7
      Top             =   6720
      Width           =   5415
   End
   Begin VB.Label l7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Treasurer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   6
      Top             =   4800
      Width           =   5355
   End
   Begin VB.Label l6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Secretary to the President"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   5
      Top             =   2880
      Width           =   5340
   End
   Begin VB.Label l5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Secretary General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   6000
      TabIndex        =   4
      Top             =   960
      Width           =   5385
   End
   Begin VB.Label l4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for External Affairs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   5385
   End
   Begin VB.Label l3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for Internal Affairs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   5415
   End
   Begin VB.Label l2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for Finance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5370
   End
   Begin VB.Label l1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "President"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   5445
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PRESIDENT (1)pres
'• VICE PRESIDENT FOR FINANCE (1)vicef
'• VICE PRESIDENT FOR INTERNAL AFFAIRS (1) viceia
'• VICE PRESIDENT FOR EXTERNAL AFFAIRS (1) viceea
'• SECRETARY GENERAL (1)secgen
'• SECRETARY TO THE PRESIDENT (1)secpres
'• TREASURER (1)tres
'• AUDITOR (1)aud
'• BOARD MEMBERS (7)board
Public Sub viewres()
Dim yearq As String
yearq = MDIForm1.lblyear.Caption
'PRESIDENT
    With Me
        .lpres.ListItems.Clear
        
    With .adowinner.Recordset
        .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l1.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lpres.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'fVICE PRESIDENT FOR FINANCE
    With Me
        .lvf.ListItems.Clear
        
    With .adowinner.Recordset
        .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l2.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lvf.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'VICE PRESIDENT FOR INTERNAL AFFAIRS
      With Me
        .lvia.ListItems.Clear
        
    With .adowinner.Recordset
         .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l3.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lvia.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'VICE PRESIDENT FOR EXTERNAL AFFAIRS
  With Me
        .lvea.ListItems.Clear
    With .adowinner.Recordset
        .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l4.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lvea.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'SECRETARY GENERAL (1)secgen
  With Me
        .lsg.ListItems.Clear
        
    With .adowinner.Recordset
         .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l5.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lsg.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'SECRETARY TO THE PRESIDENT (1)secpres
  With Me
        .lsp.ListItems.Clear
        
    With .adowinner.Recordset
         .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l6.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lsp.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'TREASURER (1)tres
  With Me
        .ltres.ListItems.Clear
        
    With .adowinner.Recordset
        .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l7.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.ltres.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'AUDITOR (1)aud
  With Me
        .lau.ListItems.Clear
        
    With .adowinner.Recordset
         .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l8.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lau.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
'BOARD MEMBERS (7)board
  With Me
        .lbm.ListItems.Clear
    With .adowinner.Recordset
         .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l9.Caption & "' and SchoolYear LIKE '" & yearq & "'"
        While Not .EOF
        Set itm = Me.lbm.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             .MoveNext
        Wend
        End With
    End With
    Me.adowinner.Refresh
End Sub


Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdexit_Click()
exitQ
End Sub

Private Sub Form_Activate()
On Error Resume Next
Call viewres
Me.adowinner.Refresh
Me.Refresh
End Sub

Private Sub l_Click(Index As Integer)

End Sub

Private Sub Label1_Click()

End Sub
