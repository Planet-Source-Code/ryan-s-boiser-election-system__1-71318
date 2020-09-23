VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form formVote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SSITE"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form3.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      Picture         =   "Form3.frx":030A
      ScaleHeight     =   780
      ScaleWidth      =   4875
      TabIndex        =   45
      Top             =   0
      Width           =   4905
   End
   Begin VB.TextBox txtlogin 
      Height          =   285
      Left            =   120
      TabIndex        =   44
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   0
      Top             =   9000
   End
   Begin VB.TextBox txttime 
      DataField       =   "StudentNo"
      DataSource      =   "adotime"
      Height          =   285
      Left            =   6960
      TabIndex        =   42
      Top             =   3840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adotime 
      Height          =   330
      Left            =   6840
      Top             =   3480
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   $"Form3.frx":8F83
      OLEDBString     =   $"Form3.frx":901D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM time_table"
      Caption         =   "adotime"
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
   Begin VB.Timer timerctr 
      Interval        =   1000
      Left            =   7440
      Top             =   5760
   End
   Begin VB.ListBox lstborig 
      Height          =   1425
      Left            =   11520
      TabIndex        =   40
      Top             =   4800
      Width           =   495
   End
   Begin VB.ListBox lboardstud 
      Height          =   3375
      Left            =   11280
      TabIndex        =   39
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox picb 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   4635
      TabIndex        =   32
      Top             =   5160
      Width           =   4695
      Begin VB.CommandButton CMDCANCELB 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton CMDOKB 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   4200
         Width           =   1935
      End
      Begin VB.ListBox LPICKB 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   2520
         TabIndex        =   34
         Top             =   240
         Width           =   1935
      End
      Begin VB.ListBox LBOARD 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Your Picked!!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2640
         TabIndex        =   36
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "List of Board Members"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear all"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      Picture         =   "Form3.frx":90B7
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtstud 
      Height          =   375
      Left            =   8760
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      DataField       =   "studentNo"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   13080
      TabIndex        =   28
      Text            =   "Text3"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12960
      Top             =   6600
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   $"Form3.frx":94F9
      OLEDBString     =   $"Form3.frx":9593
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT studentNo,Voted FROM personal_info"
      Caption         =   $"Form3.frx":962D
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
   Begin VB.TextBox Text2 
      DataField       =   "StudentNo"
      DataSource      =   "adoresult"
      Height          =   285
      Left            =   13320
      TabIndex        =   27
      Text            =   "Text2"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoresult 
      Height          =   330
      Left            =   12960
      Top             =   6960
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
      Connect         =   $"Form3.frx":96C7
      OLEDBString     =   $"Form3.frx":9761
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM candidatesinfo"
      Caption         =   $"Form3.frx":97FB
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
   Begin VB.CommandButton cmdlogout 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Picture         =   "Form3.frx":9895
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8160
      Width           =   1575
   End
   Begin VB.CommandButton cmdsubmit 
      Caption         =   "Submit!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "Form3.frx":9B9F
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "StudentNo"
      DataSource      =   "adocandidate"
      Height          =   285
      Left            =   8640
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   7
      Left            =   15240
      TabIndex        =   23
      Top             =   4560
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   6
      Left            =   15240
      TabIndex        =   22
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   15240
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   15240
      TabIndex        =   20
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   15240
      TabIndex        =   19
      Top             =   2280
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   15240
      TabIndex        =   18
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   15240
      TabIndex        =   17
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   15240
      TabIndex        =   16
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox caud 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4440
      Width           =   2415
   End
   Begin VB.ComboBox ctres 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ComboBox csp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3480
      Width           =   2415
   End
   Begin VB.ComboBox csg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ComboBox cvea 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox cvia 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ComboBox cvf 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cpres 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc adocandidate 
      Height          =   330
      Left            =   9960
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   $"Form3.frx":A469
      OLEDBString     =   $"Form3.frx":A503
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * from query_pres"
      Caption         =   $"Form3.frx":A59D
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
   Begin VB.Label lbltimenow 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   43
      Top             =   8880
      Width           =   75
   End
   Begin VB.Label lbltimer 
      Caption         =   "00 : 00 : 00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   41
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOARD MEMBERS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   1935
      TabIndex        =   31
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " AUDITOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   1200
      TabIndex        =   7
      Top             =   4440
      Width           =   810
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TREASURER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   885
      TabIndex        =   6
      Top             =   4080
      Width           =   1110
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SECRETARY TO THE PRESIDENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   420
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Width           =   1725
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SECRETARY GENERAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   1860
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VICE PRESIDENT FOR EXTERNAL AFFAIRS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2520
      Width           =   1680
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VICE PRESIDENT FOR INTERNAL AFFAIRS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1740
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VICE PRESIDENT FOR FINANCE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRESIDENT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   210
      Left            =   1185
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "formVote"
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
Dim ctr, ctrmin, ctrhour, ctr1 As Integer
Public Sub viewAll()
    With Me.adocandidate.Recordset
        'for president
        Me.adocandidate.Refresh
        .Filter = "code LIKE '" & "pres" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.cpres.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(0).AddItem .Fields("studentno")
        .MoveNext
        Loop
       'for vice for finance
         .Filter = "code LIKE '" & "vicef" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.cvf.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(1).AddItem .Fields("studentno")
        .MoveNext
        Loop
        'for vice for internal
         .Filter = "code LIKE '" & "viceia" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.cvia.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(2).AddItem .Fields("studentno")
        .MoveNext
        Loop
        'for vice for external
         .Filter = "code LIKE '" & "viceea" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.cvea.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(3).AddItem .Fields("studentno")
        .MoveNext
        Loop
    '• SECRETARY GENERAL (1)secgen
        .Filter = "code LIKE '" & "secgen" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.csg.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(4).AddItem .Fields("studentno")
        .MoveNext
        Loop
        '• SECRETARY TO THE PRESIDENT (1)secpres
        .Filter = "code LIKE '" & "secpres" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.csp.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(5).AddItem .Fields("studentno")
        .MoveNext
        Loop
        '• TREASURER (1)tres
         .Filter = "code LIKE '" & "tres" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.ctres.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(6).AddItem .Fields("studentno")
        .MoveNext
        Loop
        '• AUDITOR (1)aud
        .Filter = "code LIKE '" & "aud" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.caud.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.Combo1(7).AddItem .Fields("studentno")
        .MoveNext
        Loop
        '• BOARD MEMBERS (7)board
        .Filter = "code LIKE '" & "board" & "' and SchoolYear LIKE '" & MDIForm1.lblyear.Caption & "'"
        Do Until .EOF
        Me.LBOARD.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.lboardstud.AddItem .Fields("studentno")
        .MoveNext
        Loop
    End With
    Me.adocandidate.Refresh
End Sub

Private Sub DataCombo7_Click(Area As Integer)

End Sub

Private Sub caud_Change()
With Me
    .Combo1(7).ListIndex = .caud.ListIndex
End With
End Sub

Private Sub caud_Click()
caud_Change
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdboard_Click()
'Me.lstborig.Clear
'Me.LPICKB.Clear
'Me.lboardstud.Clear
'Me.LBOARD.Clear
' With Me.adocandidate.Recordset
' .Filter = "code LIKE '" & "board" & "'"
'        Do Until .EOF
'        Me.LBOARD.AddItem .Fields("surname") & "," & .Fields("firstname")
'        Me.lboardstud.AddItem .Fields("studentno")
'        .MoveNext
'        Loop
'    End With
'    Me.adocandidate.Refresh
'Me.picb.Visible = True
End Sub

Private Sub CMDCANCELB_Click()
Me.lstborig.Clear
Me.LPICKB.Clear
Me.lboardstud.Clear
Me.LBOARD.Clear
 With Me.adocandidate.Recordset
 .Filter = "code LIKE '" & "board" & "'"
        Do Until .EOF
        Me.LBOARD.AddItem .Fields("surname") & "," & .Fields("firstname")
        Me.lboardstud.AddItem .Fields("studentno")
        .MoveNext
        Loop
    End With
    Me.adocandidate.Refresh
Me.picb.Visible = False
End Sub

Private Sub cmdclear_Click()
With Me
    For i = 0 To 7
    .Combo1(i).Text = ""
    Next i
    Me.lstborig.Clear
    Me.LPICKB.Clear
    Me.lboardstud.Clear
    Me.LBOARD.Clear
    .cpres.Clear
    .csg.Clear
    .caud.Clear
    .csp.Clear
    .ctres.Clear
    .cvea.Clear
    .cvf.Clear
    .cvia.Clear
    Call viewAll
End With
End Sub

Private Sub cmdlogout_Click()
Form1.Show
Unload Me
End Sub

Private Sub CMDOKB_Click()
Me.picb.Visible = False
End Sub

Private Sub cmdsubmit_Click()
Dim i, j As Integer
Dim nullme As Integer
nullme = 0
For i = 0 To 7
    If Me.Combo1(i).Text <> "" Then
            nullme = nullme + 1
    End If
Next i
    If nullme < 1 And Me.LPICKB.ListCount = 0 Then
        MsgBox "Please Vote atleast '1' candidate!!!", vbOKOnly + vbExclamation, "SSITE"
        Me.Adodc1.Refresh
        Exit Sub
    End If
        Me.adoresult.Refresh
        Me.adoresult.Recordset.MoveFirst
        With Me.adoresult.Recordset
            Do Until .EOF
               
                For j = 0 To 7
                    If Me.Combo1(j).Text <> "" Then
                     .Filter = 0
                     .Filter = "StudentNo LIKE '" & Me.Combo1(j) & "'"
                    .Fields("voteCount") = .Fields("voteCount").Value + 1
                    .Update
                    End If
                Next j
               .MoveNext
            Loop
            Me.adoresult.Refresh
        End With
    'for board members
     If Me.LPICKB.ListCount <> 0 Then
         Me.adoresult.Refresh
         With Me.adoresult.Recordset
            .MoveFirst
            For i = 0 To Me.LPICKB.ListCount - 1
                .Filter = 0
                .Filter = "StudentNo LIKE '" & Me.lstborig.List(i) & "'"
                .Fields("voteCount") = .Fields("voteCount").Value + 1
                .Update
                '.MoveNext
            Next i
          End With
    End If
'=======================
    Me.adoresult.Refresh
    Me.Adodc1.Refresh
    Me.Adodc1.Recordset.Filter = "studentNo LIKE '" & Me.txtstud.Text & "'"
    Me.Adodc1.Recordset.Fields("Voted") = True
    Me.Adodc1.Recordset.Update
    '-------time to yun!!!
    With formtimer.adotime.Recordset
        ctr1 = ctr + (ctrmin * 60) + (ctrhour * 3600)
        .AddNew
        .Fields("login") = Format$(Me.txtlogin.Text, "HH:MM:SS")
        .Fields("StudentNo") = Me.txtstud.Text
        .Fields("logout") = Format$(Time, "HH:MM:SS")
        .Fields("DateVote") = Format(Now, "MM/DD/YYYY")
        .Fields("Timeconsume") = ctr1
        .Fields("SchoolYear") = MDIForm1.lblyear.Caption
        .Update
    End With
    '------------------------
    Me.Adodc1.Refresh
    Me.adoresult.Refresh
    Me.adoresult.Refresh
    Me.adoresult.Refresh
    Me.Refresh
    MsgBox "Thank you for participating the election....", vbOKOnly + vbInformation, "SSITE"
        Form1.Show
        Unload Me
End Sub

Private Sub cpres_Change()
With Me
    .Combo1(0).ListIndex = .cpres.ListIndex
End With
End Sub

Private Sub cpres_Click()
With Me
    .Combo1(0).ListIndex = .cpres.ListIndex
End With
End Sub

Private Sub csg_Change()
With Me
    .Combo1(4).ListIndex = .csg.ListIndex
End With
End Sub

Private Sub csg_Click()
csg_Change
End Sub

Private Sub csp_Change()
With Me
    .Combo1(5).ListIndex = .csp.ListIndex
End With
End Sub

Private Sub csp_Click()
csp_Change
End Sub

Private Sub ctres_Change()
With Me
    .Combo1(6).ListIndex = .ctres.ListIndex
End With
End Sub

Private Sub ctres_Click()
ctres_Change
End Sub

Private Sub cvea_Change()
With Me
    .Combo1(3).ListIndex = .cvea.ListIndex
End With
End Sub

Private Sub cvea_Click()
 cvea_Change
End Sub

Private Sub cvf_Change()
With Me
    .Combo1(1).ListIndex = .cvf.ListIndex
End With
End Sub

Private Sub cvf_Click()
cvf_Change
End Sub

Private Sub cvia_Change()
With Me
    .Combo1(2).ListIndex = .cvia.ListIndex
End With
End Sub

Private Sub cvia_Click()
cvia_Change
End Sub

Private Sub Form_Activate()
'On Error Resume Next
ctrmin = 0
ctrhour = 0
ctr1 = 0
ctr = 0
Me.timerctr.Enabled = True
Me.txtlogin.Text = Format(Time, "HH:MM:SS")
viewAll
End Sub

Private Sub Form_Load()
'Me.lmove.Height = 15
End Sub

Private Sub LBOARD_Click()
Me.lboardstud.ListIndex = Me.LBOARD.ListIndex
End Sub

Private Sub LBOARD_DblClick()
If Me.LPICKB.ListCount < 7 Then
Me.LPICKB.AddItem Me.LBOARD.Text
Me.lstborig.AddItem Me.lboardstud.Text
Me.LBOARD.RemoveItem Me.LBOARD.ListIndex
Me.lboardstud.RemoveItem Me.lboardstud.ListIndex
Else
    MsgBox "Sorry you can only vote '7' Board members...", vbOKOnly + vbInformation, "SSITE"
        Exit Sub
End If
End Sub

Private Sub LPICKB_Click()
With Me
    .lstborig.ListIndex = Me.LPICKB.ListIndex
End With
End Sub

Private Sub LPICKB_DblClick()
With Me
    .LBOARD.AddItem Me.LPICKB.Text
    .lboardstud.AddItem Me.lstborig.Text
    .LPICKB.RemoveItem Me.LPICKB.ListIndex
    .lstborig.RemoveItem Me.lstborig.ListIndex
End With
End Sub



Private Sub Timer3_Timer()
Me.Timer3.Enabled = True
    Me.lbltimenow.Caption = Format(Time, "HH:MM:SS")
End Sub

Private Sub timerctr_Timer()
Me.timerctr.Enabled = True
ctr = ctr + 1

If ctr = 60 Then
    ctrmin = ctrmin + 1
    ctr = 0
End If
If ctrmin = 60 Then
    ctrhour = ctrhour + 1
    ctrmin = 0
    ctr = 0
End If
Me.lbltimer.Caption = ctrhour & " : " & ctrmin & " : " & ctr
End Sub
