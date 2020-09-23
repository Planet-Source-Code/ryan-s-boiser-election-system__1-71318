VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formreport 
   Caption         =   "Report"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MouseIcon       =   "formreport.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11010
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "formreport.frx":030A
      ScaleHeight     =   825
      ScaleWidth      =   12225
      TabIndex        =   20
      Top             =   0
      Width           =   12255
   End
   Begin VB.TextBox txttime 
      DataField       =   "StudentNo"
      DataSource      =   "adoAMPM"
      Height          =   375
      Left            =   5520
      TabIndex        =   19
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc adoAMPM 
      Height          =   375
      Left            =   1680
      Top             =   11280
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      Connect         =   $"formreport.frx":B006
      OLEDBString     =   $"formreport.frx":B0A0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM time_table"
      Caption         =   "adoAMPM"
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
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9840
      Picture         =   "formreport.frx":B13A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   8895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11775
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   9855
         Begin VB.Label lyear 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Year"
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
            Height          =   360
            Left            =   8160
            TabIndex        =   11
            Top             =   0
            Width           =   1650
         End
         Begin VB.Label lcourse 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Course"
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
            Height          =   360
            Left            =   6360
            TabIndex        =   10
            Top             =   0
            Width           =   1890
         End
         Begin VB.Label lfirst 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Firstname"
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
            Height          =   360
            Left            =   4080
            TabIndex        =   9
            Top             =   0
            Width           =   2370
         End
         Begin VB.Label lsurname 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Surname"
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
            Height          =   360
            Left            =   1560
            TabIndex        =   8
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label lstud 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student No."
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
            Height          =   360
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.TextBox txtstud 
         DataField       =   "StudentNo"
         DataSource      =   "adojoin"
         Height          =   375
         Left            =   16320
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   720
         TabIndex        =   1
         Top             =   4080
         Width           =   9855
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Year"
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
            Height          =   720
            Left            =   8160
            TabIndex        =   17
            Top             =   0
            Width           =   1650
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Course"
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
            Height          =   720
            Left            =   6360
            TabIndex        =   16
            Top             =   0
            Width           =   1890
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Firstname"
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
            Height          =   720
            Left            =   4080
            TabIndex        =   15
            Top             =   0
            Width           =   2370
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Surname"
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
            Height          =   720
            Left            =   1560
            TabIndex        =   14
            Top             =   0
            Width           =   2610
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Student No."
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
            Height          =   720
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1575
         End
      End
      Begin MSAdodcLib.Adodc adojoin 
         Height          =   375
         Left            =   14040
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         Connect         =   $"formreport.frx":BE04
         OLEDBString     =   $"formreport.frx":BE9E
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM personal_info"
         Caption         =   "adostudent"
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
      Begin MSComctlLib.ListView lvjoin 
         Height          =   2895
         Left            =   720
         TabIndex        =   5
         Top             =   600
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5106
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
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
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvkj 
         Height          =   2775
         Left            =   720
         TabIndex        =   12
         Top             =   4080
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4895
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
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
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Line Line10 
         X1              =   7440
         X2              =   7440
         Y1              =   7560
         Y2              =   8640
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6720
         TabIndex        =   34
         Top             =   7680
         Width           =   555
      End
      Begin VB.Line Line9 
         X1              =   5520
         X2              =   5520
         Y1              =   7560
         Y2              =   8640
      End
      Begin VB.Line Line8 
         X1              =   6480
         X2              =   6480
         Y1              =   7560
         Y2              =   8640
      End
      Begin VB.Line Line7 
         X1              =   5520
         X2              =   9600
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line6 
         X1              =   7560
         X2              =   7560
         Y1              =   7560
         Y2              =   8640
      End
      Begin VB.Line Line5 
         X1              =   8640
         X2              =   8640
         Y1              =   7560
         Y2              =   8640
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Percentage %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   7560
         TabIndex        =   25
         Top             =   7200
         Width           =   2055
      End
      Begin VB.Line Line4 
         X1              =   5520
         X2              =   9600
         Y1              =   8640
         Y2              =   8640
      End
      Begin VB.Line Line3 
         X1              =   9600
         X2              =   9600
         Y1              =   7320
         Y2              =   8640
      End
      Begin VB.Line Line2 
         X1              =   5520
         X2              =   9600
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Line Line1 
         X1              =   5520
         X2              =   9600
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Label l2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6720
         TabIndex        =   33
         Top             =   8400
         Width           =   555
      End
      Begin VB.Label l1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "100 %"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6720
         TabIndex        =   32
         Top             =   8040
         Width           =   555
      End
      Begin VB.Label lbldiff 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5520
         TabIndex        =   31
         Top             =   8400
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total No. of Students who not yet voted"
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
         Height          =   360
         Left            =   1560
         TabIndex        =   30
         Top             =   8280
         Width           =   3900
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblparticipate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   8040
         Width           =   975
      End
      Begin VB.Label lbltotalstud 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5520
         TabIndex        =   28
         Top             =   7680
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total No. of Students who participated"
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
         Height          =   360
         Left            =   1560
         TabIndex        =   27
         Top             =   7920
         Width           =   3900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total No. of Students"
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
         Height          =   360
         Left            =   1560
         TabIndex        =   26
         Top             =   7560
         Width           =   3900
      End
      Begin VB.Label lblPM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00.0%"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
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
         Height          =   240
         Left            =   8760
         TabIndex        =   24
         Top             =   8040
         Width           =   735
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "PM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   23
         Top             =   7560
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00.0%"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
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
         Height          =   240
         Left            =   7680
         TabIndex        =   22
         Top             =   8040
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "AM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   7560
         Width           =   1065
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "List of Students who participated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Top             =   240
         Width           =   4560
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "List of Students who did ""NOT"" participated"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   2520
         TabIndex        =   3
         Top             =   3720
         Width           =   6030
      End
   End
End
Attribute VB_Name = "formreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total, join, KJ, totalAMPM As Integer
Dim KJpercent, joinpercent, joinAM, joinPMpercent, joinPM, joinAMpercent As Double
Dim itm
Public Sub info()
    Me.adojoin.Refresh
    With Me.adojoin.Recordset
        Me.adoAMPM.Recordset.Filter = "dateVote LIKE '" & Format$(Date, "MM/DD/YYYY") & "'"
        If .RecordCount <> 0 Then
        total = .RecordCount
        .MoveFirst
        Do Until .EOF
            If .Fields("voted") = True Then
                join = join + 1
            'for those who joined
            Set itm = Me.lvjoin.ListItems.Add(, , .Fields!StudentNo)
                 itm.SubItems(1) = .Fields("Surname")
                 itm.SubItems(2) = .Fields("firstname")
                 itm.SubItems(3) = .Fields("Course")
                 itm.SubItems(4) = .Fields("Year")
            Else
                KJ = KJ + 1
                'for those who not joined
            Set itm = Me.lvkj.ListItems.Add(, , .Fields!StudentNo)
                 itm.SubItems(1) = .Fields("Surname")
                 itm.SubItems(2) = .Fields("firstname")
                 itm.SubItems(3) = .Fields("Course")
                 itm.SubItems(4) = .Fields("Year")
            End If
            .MoveNext
        Loop
            Me.adojoin.Refresh
        End If
    End With
End Sub
Public Sub AMPMStat()
    Me.adojoin.Refresh
    totalAMPM = Me.adojoin.Recordset.RecordCount
        With Me.adoAMPM.Recordset
             'total = .RecordCount
            .Filter = "DateVote LIke '" & Format$(Date, "MM/DD/YYYY") & "'"
            If .RecordCount <> 0 Then
            Do Until .EOF
                If .Fields("Logout") >= TimeValue("12.00.01") Then
                    joinPM = joinPM + 1
                ElseIf .Fields("Logout") <= TimeValue("12.00.00") Then
                    joinAM = joinAM + 1
                End If
                .MoveNext
            Loop
                
            End If
                Me.adoAMPM.Refresh
        End With
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Activate()
join = 0
total = 0
KJ = 0
joinAM = 0
joinPM = 0
Call info
Me.lbltotalstud.Caption = total
Me.lblparticipate.Caption = join
Me.lbldiff.Caption = KJ
'---over all
joinpercent = (join / total)
KJpercent = (KJ / total)
l1.Caption = Format(joinpercent, "0.00%")
l2.Caption = Format(KJpercent, "0.00%")
'---AMPM status
Call AMPMStat
joinAMpercent = joinAM / totalAMPM
joinPMpercent = joinPM / totalAMPM
Me.lblAM.Caption = Format$(joinAMpercent, "0.000%")
Me.lblPM.Caption = Format$(joinPMpercent, "0.000%")
'---PM status
End Sub

Private Sub Form_Load()
With Me.lvjoin
    .ColumnHeaders(1).Width = Me.lstud.Width
    .ColumnHeaders(2).Width = Me.lsurname.Width
    .ColumnHeaders(3).Width = Me.lfirst.Width
    .ColumnHeaders(4).Width = Me.lcourse.Width
    .ColumnHeaders(5).Width = Me.lyear.Width
End With
With Me.lvkj
    .ColumnHeaders(1).Width = Me.lstud.Width
    .ColumnHeaders(2).Width = Me.lsurname.Width
    .ColumnHeaders(3).Width = Me.lfirst.Width
    .ColumnHeaders(4).Width = Me.lcourse.Width
    .ColumnHeaders(5).Width = Me.lyear.Width
End With
End Sub

Private Sub Label17_Click()

End Sub

