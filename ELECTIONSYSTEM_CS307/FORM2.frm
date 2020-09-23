VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form formstudent 
   Caption         =   "SSITE"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MouseIcon       =   "FORM2.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   13425
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "FORM2.frx":030A
      ScaleHeight     =   825
      ScaleWidth      =   12225
      TabIndex        =   57
      Top             =   0
      Width           =   12255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sort Option"
      Height          =   735
      Left            =   240
      TabIndex        =   55
      Top             =   9000
      Width           =   7215
      Begin VB.CommandButton cmdgo 
         Caption         =   "Go!"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optstud 
         Caption         =   "Student No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optsurname 
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox cbosort 
         Height          =   360
         ItemData        =   "FORM2.frx":B006
         Left            =   795
         List            =   "FORM2.frx":B013
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   450
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "FORM2.frx":B034
      Height          =   3855
      Left            =   240
      TabIndex        =   54
      Top             =   5160
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   6800
      _Version        =   393216
      ForeColor       =   0
      Cols            =   12
      FixedCols       =   0
      ForeColorFixed  =   -2147483647
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      MergeCells      =   4
      AllowUserResizing=   3
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0)._NumMapCols=   12
      _Band(0)._MapCol(0)._Name=   "StudentNo"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Firstname"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "Surname"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "MiddleInitial"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "Course"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "Year"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "Gender"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "Civil_Status"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "Nationality"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "Special_Password"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "Voted"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(11)._Name=   "Date_birth"
      _Band(0)._MapCol(11)._RSIndex=   11
   End
   Begin VB.ComboBox cbostatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FORM2.frx":B052
      Left            =   2160
      List            =   "FORM2.frx":B05C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   30
      Top             =   720
      Width           =   10935
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   4800
         TabIndex        =   53
         Top             =   2400
         Width           =   6015
         Begin VB.TextBox txtfind 
            Height          =   375
            Left            =   1680
            TabIndex        =   23
            Top             =   1320
            Width           =   2295
         End
         Begin VB.CommandButton Cmdfind 
            Caption         =   "Find by Grid"
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
            Left            =   240
            TabIndex        =   24
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Timer Timer1 
            Interval        =   20
            Left            =   0
            Top             =   1320
         End
         Begin VB.CommandButton cmdexit 
            Caption         =   "Close"
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
            Left            =   4440
            Picture         =   "FORM2.frx":B071
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "Add"
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
            Picture         =   "FORM2.frx":BD3B
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdedit 
            Caption         =   "Edit"
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
            Left            =   3000
            Picture         =   "FORM2.frx":C4A5
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmdsave 
            Caption         =   "Save"
            Enabled         =   0   'False
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
            Left            =   1560
            Picture         =   "FORM2.frx":C8E7
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   120
            Width           =   1455
         End
         Begin VB.CommandButton cmddel 
            Caption         =   "Delete"
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
            Left            =   4440
            Picture         =   "FORM2.frx":D5B1
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox txtpass 
         DataField       =   "Special_Password"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CommandButton c4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10320
         Picture         =   "FORM2.frx":D9F3
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton c3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9840
         Picture         =   "FORM2.frx":DE35
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton c2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         Picture         =   "FORM2.frx":E277
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton c1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         Picture         =   "FORM2.frx":E6B9
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox cboyear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FORM2.frx":EAFB
         Left            =   1920
         List            =   "FORM2.frx":EAFD
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cbocourse 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FORM2.frx":EAFF
         Left            =   1920
         List            =   "FORM2.frx":EB0C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtyear 
         DataField       =   "Year"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtcourse 
         DataField       =   "Course"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ComboBox cbosex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FORM2.frx":EB21
         Left            =   1920
         List            =   "FORM2.frx":EB2B
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtbday 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53739521
         CurrentDate     =   37987
      End
      Begin VB.TextBox txtstudNo 
         DataField       =   "StudentNo"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtfname 
         DataField       =   "Firstname"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtsname 
         DataField       =   "Surname"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtmname 
         DataField       =   "MiddleInitial"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtbirth 
         DataField       =   "Date_birth"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtsex 
         DataField       =   "Gender"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtcstat 
         DataField       =   "Civil_Status"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtNa 
         DataField       =   "Nationality"
         DataSource      =   "AdopersonalINfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   4800
         TabIndex        =   31
         Top             =   240
         Width           =   6015
         Begin VB.TextBox txtpri 
            DataField       =   "P_Address"
            DataSource      =   "Adoaddress"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   4935
         End
         Begin VB.TextBox txtsec 
            DataField       =   "S_Address"
            DataSource      =   "Adoaddress"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtter 
            DataField       =   "T_Address"
            DataSource      =   "Adoaddress"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Primary"
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
            Left            =   360
            TabIndex        =   35
            Top             =   480
            Width           =   630
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Secondary"
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
            Left            =   90
            TabIndex        =   34
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tertiary"
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
            Left            =   225
            TabIndex        =   33
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H80000001&
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   6135
         End
      End
      Begin MSAdodcLib.Adodc AdopersonalINfo 
         Height          =   330
         Left            =   1440
         Top             =   7200
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
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
         Connect         =   $"FORM2.frx":EB3D
         OLEDBString     =   $"FORM2.frx":EBD7
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM Personal_Info"
         Caption         =   "AdopersonalINfo"
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
      Begin VB.Label lblpass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lbinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SSITE ELECTION SYSTEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   5040
         TabIndex        =   51
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Course:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Student No.:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname:"
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
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Initial"
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
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
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
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Stat"
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
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2760
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adoaddress 
      Height          =   330
      Left            =   15600
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"FORM2.frx":EC71
      OLEDBString     =   $"FORM2.frx":ED0B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Address_Info"
      Caption         =   "address"
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
   Begin VB.TextBox txtaddressID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13320
      TabIndex        =   44
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "formstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addressnav()
    With Me.Adoaddress.Recordset
       ' .Filter = 0
        .Filter = "StudentNo LIKE '" & Me.txtstudNo.Text & "'"
        'Me.Adoaddress.Refresh
    End With
End Sub
Public Sub unlockedme()
With Me
    .txtstudNo.Locked = False
    .txtfname.Locked = False
    .txtsname.Locked = False
    .txtmname.Locked = False
    .dtbday.Visible = True
    .cbocourse.Visible = True
    .cbosex.Visible = True
    .cbostatus.Visible = True
    .cboyear.Visible = True
    .txtpass.Locked = False
    .txtpri.Locked = False
    .txtsec.Locked = False
    .txtter.Locked = False
    .cmdadd.Enabled = False
    .cmdsave.Enabled = True
    .txtNa.Locked = False
    .cmddel.Caption = "Cancel"
    .cmdedit.Enabled = False
    .txtsex.Visible = False
    .txtcstat.Visible = False
    .txtyear.Visible = False
    .txtbirth.Visible = False
    .txtcourse.Visible = False
    .c1.Enabled = False
    .c2.Enabled = False
    .c3.Enabled = False
    .c4.Enabled = False
  '  .cbocourse.List (1)
  '  .cbosex.List (1)
End With
End Sub
Public Sub lockedme()
With Me
    .txtstudNo.Locked = True
    .txtfname.Locked = True
    .txtsname.Locked = True
    .txtmname.Locked = True
    .dtbday.Visible = False
    .txtNa.Locked = True
    .cbocourse.Visible = False
    .cbosex.Visible = False
    .cbostatus.Visible = False
    .cboyear.Visible = False
    .txtpass.Locked = True
    .txtpri.Locked = True
    .txtsec.Locked = True
    .txtter.Locked = True
    .cmdadd.Enabled = True
    .cmdedit.Enabled = True
    .cmdsave.Enabled = False
    .cmddel.Caption = "Delete"
    .txtsex.Visible = True
    .txtcstat.Visible = True
    .txtyear.Visible = True
    .txtbirth.Visible = True
    .txtcourse.Visible = True
    .c1.Enabled = True
    .c2.Enabled = True
    .c3.Enabled = True
    .c4.Enabled = True
End With
End Sub


Private Sub c1_Click()
Me.AdopersonalINfo.Recordset.MoveFirst
Call addressnav
End Sub

Private Sub c2_Click()
With Me.AdopersonalINfo.Recordset
    .MovePrevious
    If .BOF Then .MoveFirst
End With
Call addressnav
End Sub

Private Sub c3_Click()
With Me.AdopersonalINfo.Recordset
    .MoveNext
    If .EOF Then .MoveLast
End With
Call addressnav
End Sub

Private Sub c4_Click()

Me.AdopersonalINfo.Recordset.MoveLast
Call addressnav
End Sub

Private Sub cbocourse_Change()
Me.cboyear.Clear
Me.txtcourse.Text = Me.cbocourse.Text
With Me.cboyear
If Me.cbocourse.Text = "ACT" Then
        .AddItem "1st"
        .AddItem "2nd"
ElseIf Me.cbocourse.Text = "BSIT" Or Me.cbocourse.Text = "BSCS" Then
        .AddItem "1st"
        .AddItem "2nd"
        .AddItem "3rd"
        .AddItem "4th"
End If
 End With
End Sub

Private Sub cbocourse_Click()
Me.cboyear.Clear
Me.txtcourse.Text = Me.cbocourse.Text
With Me.cboyear
If Me.cbocourse.Text = "ACT" Then
        .AddItem "1st"
        .AddItem "2nd"
ElseIf Me.cbocourse.Text = "BSIT" Or Me.cbocourse.Text = "BSCS" Then
        .AddItem "1st"
        .AddItem "2nd"
        .AddItem "3rd"
        .AddItem "4th"
End If
 End With
 End Sub

Private Sub cbosex_Change()
Me.txtsex.Text = Me.cbosex.Text
End Sub

Private Sub cbosex_Click()
Me.txtsex.Text = Me.cbosex.Text
End Sub

Private Sub cbostatus_Change()
Me.txtcstat.Text = Me.cbostatus.Text
End Sub

Private Sub cbostatus_Click()
Me.txtcstat.Text = Me.cbostatus.Text
End Sub

Private Sub cboyear_Change()
Me.txtyear.Text = Me.cboyear.Text
End Sub

Private Sub cboyear_Click()
Me.txtyear.Text = Me.cboyear.Text
End Sub

Private Sub cmdadd_Click()
With Me
    Call unlockedme
    .txtstudNo.SetFocus
    .AdopersonalINfo.Recordset.AddNew
    .Adoaddress.Recordset.AddNew
    .cmdedit.Enabled = False
End With
End Sub


Private Sub cmddel_Click()
Dim askmedel As Single
Dim mydelete As String
With Me
    If .cmddel.Caption = "Delete" Then
        askmedel = MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion, "SSITE")
            If askmedel = vbYes Then
                mydelete = .AdopersonalINfo.Recordset.Fields("StudentNo")
                .AdopersonalINfo.Recordset.Delete
                .AdopersonalINfo.Recordset.MoveNext
                If .AdopersonalINfo.Recordset.EOF Then: .AdopersonalINfo.Refresh
                .Adoaddress.Recordset.Delete
                .Adoaddress.Recordset.MoveNext
                If .Adoaddress.Recordset.EOF Then: .Adoaddress.Refresh
                With formVote.adoresult.Recordset
                    .Filter = "StudentNo LIKE '" & mydelete & "'"
                    If .RecordCount <> 0 Then
                        .Delete
                        .MoveNext
                    End If
                End With
            End If
                .Adoaddress.Refresh
                .AdopersonalINfo.Refresh
                .Adoaddress.Refresh
                .AdopersonalINfo.Refresh
                .Refresh
    Else
        On Error Resume Next
        .Adoaddress.Recordset.CancelUpdate
        .AdopersonalINfo.Recordset.CancelUpdate
        Call lockedme
        .Refresh
        .Refresh
    End If
End With
End Sub

Private Sub cmdedit_Click()
With Me
    Call unlockedme
End With
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub Cmdfind_Click()
Dim i, j As Integer
If Me.txtfind.Text <> "" Then
With Me.MSHFlexGrid1
.FillStyle = flexFillRepeat
.Col = 0
.Row = 0
.ColSel = .Cols - 1
.RowSel = .Rows - 1
.CellFontBold = False
.FillStyle = flexFillSingle
For i = 0 To .Cols - 1
    For j = 1 To .Rows - 1
        If InStr(UCase(.TextMatrix(j, i)), UCase(Me.txtfind.Text)) Then
            .Col = i
            .Row = j
            .CellFontBold = True
        End If
    Next j
Next i
End With
End If
End Sub

Private Sub cmdgo_Click()
With Me
    If .cbosort.ListIndex = 1 Then
            If .optsurname.Value Then
                .AdopersonalINfo.Recordset.Sort = "Surname ASC"
            ElseIf .optstud.Value = True Then
                .AdopersonalINfo.Recordset.Sort = "StudentNo ASC"
            End If
    ElseIf .cbosort.ListIndex = 2 Then
            If .optsurname.Value Then
                .AdopersonalINfo.Recordset.Sort = "Surname DESC"
            ElseIf .optstud.Value = True Then
                .AdopersonalINfo.Recordset.Sort = "StudentNo DESC"
            End If
    End If
End With
End Sub

Private Sub cmdsave_Click()
With Me
    If .txtstudNo.Text <> "" And .txtfname.Text <> "" _
        And .txtsname.Text <> "" And .txtmname.Text <> "" _
        And .txtcourse.Text <> "" And .txtyear.Text <> "" _
        And .txtsex.Text <> "" And .txtcstat.Text <> "" _
        And .txtNa.Text <> "" And .txtpass.Text <> "" _
        And .txtbirth.Text <> "" And .txtpri.Text <> "" Then
        .AdopersonalINfo.Recordset.Fields("StudentNo") = .txtstudNo.Text
        .AdopersonalINfo.Recordset.Fields("Firstname") = .txtfname.Text
        .AdopersonalINfo.Recordset.Fields("Surname") = .txtsname.Text
        .AdopersonalINfo.Recordset.Fields("MiddleInitial") = .txtmname.Text
        .AdopersonalINfo.Recordset.Fields("Course") = .txtcourse.Text
        .AdopersonalINfo.Recordset.Fields("Year") = .txtyear.Text
        .AdopersonalINfo.Recordset.Fields("Gender") = .txtsex.Text
        .AdopersonalINfo.Recordset.Fields("Civil_Status") = .txtcstat.Text
        .AdopersonalINfo.Recordset.Fields("Nationality") = .txtNa.Text
        .AdopersonalINfo.Recordset.Fields("Special_Password") = .txtpass.Text
        .AdopersonalINfo.Recordset.Fields("Date_birth") = .txtbirth.Text
        .AdopersonalINfo.Recordset.Update
        .Adoaddress.Recordset.Fields("StudentNo") = .txtstudNo.Text
        .Adoaddress.Recordset.Fields("P_Address") = .txtpri.Text
        If .txtsec.Text <> "" Then
        .Adoaddress.Recordset.Fields("S_Address") = .txtsec.Text
        End If
        If .txtter.Text <> "" Then
        .Adoaddress.Recordset.Fields("T_Address") = .txtter.Text
        End If
        .Adoaddress.Recordset.Update
        .Adoaddress.Refresh
        .AdopersonalINfo.Refresh
        .Adoaddress.Refresh
        .AdopersonalINfo.Refresh
        Me.Refresh
        Call lockedme
    Else
        MsgBox "Please fill up all requirements...", vbOKOnly + vbInformation, "SSITE"
            Me.txtstudNo.SetFocus
            Exit Sub
    End If
End With
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtbday_Change()
Me.txtbirth.Text = Format$(Me.dtbday.Value, "MMMM DD, YYYY")
End Sub

Private Sub dtbday_Click()
Me.txtbirth.Text = Format$(Me.dtbday.Value, "MMMM DD, YYYY")
End Sub

Private Sub Form_Activate()
Me.WindowState = 2
'Me.Adoaddress.Refresh
'Me.AdopersonalINfo.Refresh
Me.dtbday.Value = Format$(Date, "MMMM DD, YYYY")
Me.txtbirth.Text = Format$(Date, "MMMM DD, YYYY")
Me.cbosort.ListIndex = 0
Call addressnav
End Sub

Private Sub txtfind_Change()
Cmdfind_Click
End Sub

Private Sub txtfind_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyBack Then
    Cmdfind_Click
End If
End Sub

Private Sub txtfind_KeyUp(KeyCode As Integer, Shift As Integer)
Cmdfind_Click
End Sub

Private Sub txtstudNo_Change()
Me.txtaddressID.Text = Me.txtstudNo.Text
End Sub

Private Sub txtstudNo_LostFocus()
Me.txtaddressID.Text = Me.txtstudNo.Text
End Sub
