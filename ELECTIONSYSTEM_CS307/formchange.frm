VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form formchange 
   BorderStyle     =   0  'None
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "formchange.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000001&
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txt1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txt3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000001&
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton cmdapply 
         Caption         =   "Apply"
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
         Left            =   240
         Picture         =   "formchange.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
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
         Left            =   2280
         Picture         =   "formchange.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   600
         Picture         =   "formchange.frx":189E
         ScaleHeight     =   1215
         ScaleWidth      =   3750
         TabIndex        =   2
         Top             =   120
         Width           =   3780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NewPassword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ConfirmPassword"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   1695
      End
   End
   Begin VB.TextBox txtadmin 
      DataField       =   "passwordADmin"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   7440
      Width           =   3375
      _ExtentX        =   5953
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
      Connect         =   $"formchange.frx":928D
      OLEDBString     =   $"formchange.frx":9327
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Admin_tbl"
      Caption         =   "Adodc1"
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
Attribute VB_Name = "formchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
With Me
    If .Check1.Value = vbChecked Then
        .txt1.PasswordChar = ""
        .txt2.PasswordChar = ""
        .txt3.PasswordChar = ""
    Else
        .txt1.PasswordChar = "*"
        .txt2.PasswordChar = "*"
        .txt3.PasswordChar = "*"
    End If
End With
End Sub

Private Sub cmdapply_Click()
With Me
    If .txt1.Text = .Adodc1.Recordset.Fields("passwordADmin") Then
        If .txt3.Text = .txt2.Text Then
            .Adodc1.Recordset.Fields("passwordADmin") = .txt2.Text
            .Adodc1.Recordset.Update
            .Adodc1.Refresh
            MsgBox "Password successfully change!!!", vbOKOnly + vbMsgBoxRtlReading, "SSITE"
            Unload Me
            Exit Sub
        Else
            MsgBox "Sorry not match...Please checked each value", vbOKOnly + vbExclamation, "SSITE"
                Exit Sub
        End If
    Else
        MsgBox "Password not match", vbOKOnly + vbExclamation, "SSITE"
            Exit Sub
    End If
End With
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Move 0, 0
End Sub

