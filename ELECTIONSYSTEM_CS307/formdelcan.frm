VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form formyear 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "formdelcan.frx":0000
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   3360
      ScaleHeight     =   3825
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   2280
      Width           =   3855
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox cboyear 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Text            =   "cboyear"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   0
         Picture         =   "formdelcan.frx":030A
         ScaleHeight     =   1275
         ScaleWidth      =   3795
         TabIndex        =   2
         Top             =   0
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change School Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   3135
      End
   End
   Begin VB.TextBox txtstud 
      DataField       =   "StudentNo"
      DataSource      =   "adocandidate"
      Height          =   285
      Left            =   10320
      TabIndex        =   0
      Top             =   7920
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
      Connect         =   $"formdelcan.frx":7CF9
      OLEDBString     =   $"formdelcan.frx":7D93
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
End
Attribute VB_Name = "formyear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
If Me.cboyear.Text <> "" Then
MDIForm1.lblyear.Caption = Me.cboyear.Text
Unload Me
End If
End Sub

Private Sub Form_Activate()
Me.Move 0, 0
Me.cboyear.Text = Format$(Date, "YYYY")
End Sub

Private Sub Form_Load()
For i = 1925 To 9999
    Me.cboyear.AddItem Str(i)
Next i
End Sub

