VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formwinner 
   Caption         =   "Winner"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   MouseIcon       =   "formwinners.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      Picture         =   "formwinners.frx":030A
      ScaleHeight     =   915
      ScaleWidth      =   15360
      TabIndex        =   30
      Top             =   0
      Width           =   15390
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9720
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
      Left            =   8880
      Picture         =   "formwinners.frx":109DC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9720
      Width           =   2535
   End
   Begin VB.TextBox txtstud 
      DataField       =   "Surname"
      DataSource      =   "adowinner"
      Height          =   285
      Left            =   18000
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adowinner 
      Height          =   330
      Left            =   17880
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   $"formwinners.frx":116A6
      OLEDBString     =   $"formwinners.frx":11740
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * from res_query"
      Caption         =   "adowinner"
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
   Begin VB.Timer Timeall 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7080
      Top             =   2400
   End
   Begin MSComctlLib.ListView lpres 
      Height          =   1335
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
   Begin MSComctlLib.ListView lvf 
      Height          =   1335
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   2
      Top             =   4800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   1560
      TabIndex        =   3
      Top             =   6480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   8520
      TabIndex        =   4
      Top             =   1440
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   8520
      TabIndex        =   5
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   8520
      TabIndex        =   6
      Top             =   4800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   1335
      Left            =   8520
      TabIndex        =   7
      Top             =   6480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2355
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
      Height          =   2415
      Left            =   1560
      TabIndex        =   8
      Top             =   8280
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
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
   Begin VB.Label lblaud 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   14040
      TabIndex        =   26
      Top             =   6600
      Visible         =   0   'False
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbltres 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   14040
      TabIndex        =   25
      Top             =   4920
      Visible         =   0   'False
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblsp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   14010
      TabIndex        =   24
      Top             =   3240
      Visible         =   0   'False
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   14010
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblboard 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   7080
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblvea 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   7080
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblvia 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   7080
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblvf 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   7035
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblpres 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "There is a tie here!!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   7065
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1185
      WordWrap        =   -1  'True
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "President"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   17
      Top             =   1080
      Width           =   5445
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for Finance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   16
      Top             =   2760
      Width           =   5445
   End
   Begin VB.Label l3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for Internal Affairs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   15
      Top             =   4440
      Width           =   5415
   End
   Begin VB.Label l4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vice president for External Affairs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   14
      Top             =   6120
      Width           =   5385
   End
   Begin VB.Label l5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Secretary General"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8520
      TabIndex        =   13
      Top             =   1080
      Width           =   5445
   End
   Begin VB.Label l6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Secretary to the President"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8520
      TabIndex        =   12
      Top             =   2760
      Width           =   5445
   End
   Begin VB.Label l7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Treasurer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8520
      TabIndex        =   11
      Top             =   4440
      Width           =   5445
   End
   Begin VB.Label l8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auditor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   8520
      TabIndex        =   10
      Top             =   6120
      Width           =   5445
   End
   Begin VB.Label l9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Board Members"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   7920
      Width           =   5385
   End
End
Attribute VB_Name = "formwinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TopPres, Topvf, Topvia, Topvea, TopSg, topSp, TopTres, TopAud, Topboard, boardctr As Integer
Dim presprob, vfprob, viaprob, veaprob, sgprob, spprob, tresprob, audprob, boardprob As Boolean

Public Sub viewres()
Me.adowinner.Refresh
On Error Resume Next
'PRESIDENT
    With Me
        .lpres.ListItems.Clear
        
    With .adowinner.Recordset
        .Sort = "VoteCount DESC"
        .Filter = "Title LIKE '" & l1.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        TopPres = .Fields("VoteCount").Value
        While Not .EOF
            If TopPres = .Fields("VoteCount").Value Then
            Set itm = Me.lpres.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l2.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
         Topvf = .Fields("VoteCount").Value
        While Not .EOF
         If Topvf = .Fields("VoteCount").Value Then
            Set itm = Me.lvf.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l3.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Topvia = .Fields("VoteCount").Value
        While Not .EOF
         If Topvia = .Fields("VoteCount").Value Then
            Set itm = Me.lvia.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l4.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
         Topvea = .Fields("VoteCount").Value
        While Not .EOF
         If Topvea = .Fields("VoteCount").Value Then
            Set itm = Me.lvea.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l5.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
       TopSg = .Fields("VoteCount").Value
        While Not .EOF
         If TopSg = .Fields("VoteCount").Value Then
            Set itm = Me.lsg.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l6.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        topSp = .Fields("VoteCount").Value
        While Not .EOF
         If topSp = .Fields("VoteCount").Value Then
            Set itm = Me.lsp.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l7.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        TopTres = .Fields("VoteCount").Value
        While Not .EOF
         If TopTres = .Fields("VoteCount").Value Then
            Set itm = Me.ltres.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l8.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Topvf = .Fields("VoteCount").Value
        While Not .EOF
         If TopAud = .Fields("VoteCount").Value Then
            Set itm = Me.lau.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
             End If
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
        .Filter = "Title LIKE '" & l9.Caption & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Topboard = .Fields("VoteCount").Value
        boardctr = 0
        Do Until .EOF Or boardctr = 7
            'If boardctr <= 7 Then
            If Topboard = .Fields("VoteCount").Value Then
          'boardctr = boardtr + 1
             Set itm = Me.lbm.ListItems.Add(, , .Fields!Firstname)
             itm.SubItems(1) = .Fields("Surname")
             itm.SubItems(2) = .Fields("Course")
             itm.SubItems(3) = .Fields("VoteCount")
            Else
                
                Topboard = .Fields("VoteCount").Value
                Set itm = Me.lbm.ListItems.Add(, , .Fields!Firstname)
                itm.SubItems(1) = .Fields("Surname")
                itm.SubItems(2) = .Fields("Course")
                itm.SubItems(3) = .Fields("VoteCount")
            End If
             boardctr = boardctr + 1
             .MoveNext
            ' Else
             '   Exit Do
           ' End If
        Loop
        End With
    End With
    'vfprob, viaprob, veaprob, sgprob, spprob, tresprob, audprob, boardprob As Boolean
    With Me
    If .lpres.ListItems.Count > 1 Then presprob = True
    If .lvf.ListItems.Count > 1 Then vfprob = True
    If .lvia.ListItems.Count > 1 Then viaprob = True
    If .lvea.ListItems.Count > 1 Then veaprob = True
    If .lsg.ListItems.Count > 1 Then sgprob = True
    If .lsp.ListItems.Count > 1 Then spprob = True
    If .ltres.ListItems.Count > 1 Then tresprob = True
    If .lau.ListItems.Count > 1 Then audprob = True
    If .lbm.ListItems.Count > 7 Then boardprob = True
    Timeall.Enabled = True
    End With
    Me.adowinner.Refresh
End Sub

Private Sub cmdclose_Click()
MDIForm1.Show
Unload Me
End Sub

Private Sub cmdexit_Click()
exitQ
End Sub
Private Sub Form_Activate()
boardctr = 0
TopPres = 0
Topvf = 0
Topvia = 0
Topvea = 0
TopSg = 0
topSp = 0
TopTres = 0
TopAud = 0
Topboard = 0
Call viewres
End Sub

Private Sub Timeall_Timer()
Me.Timeall.Enabled = True
If presprob = True Then Me.lblpres.Visible = Not Me.lblpres.Visible
If vfprob = True Then Me.lblvf.Visible = Not Me.lblvf.Visible
If viaprob = True Then Me.lblvia.Visible = Not Me.lblvia.Visible
If veaprob = True Then Me.lblvea.Visible = Not Me.lblvea.Visible
If sgprob = True Then Me.lblsg.Visible = Not Me.lblsg.Visible
If spprob = True Then Me.lblsp.Visible = Not Me.lblsp.Visible
If tresprob = True Then Me.lbltres.Visible = Not Me.lbltres.Visible
If audprob = True Then Me.lblaud.Visible = Not Me.lblaud.Visible
If boardprob = True Then Me.lblboard.Visible = Not Me.lblboard.Visible
End Sub



