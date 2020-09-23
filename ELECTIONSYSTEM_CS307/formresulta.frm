VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formapply 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "formresulta.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7500
      Left            =   1320
      Picture         =   "formresulta.frx":030A
      ScaleHeight     =   7500
      ScaleWidth      =   5205
      TabIndex        =   9
      Top             =   1080
      Width           =   5205
      Begin VB.ComboBox cbostudno 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cbopos 
         Height          =   315
         ItemData        =   "formresulta.frx":118CB
         Left            =   1800
         List            =   "formresulta.frx":118EA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "Submit"
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
         Left            =   240
         Picture         =   "formresulta.frx":119AF
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   2175
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
         Left            =   2400
         Picture         =   "formresulta.frx":12279
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4440
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20709377
         CurrentDate     =   37987
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   480
         TabIndex        =   17
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   720
         TabIndex        =   16
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   360
         TabIndex        =   15
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname, Firsname"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   3240
         TabIndex        =   13
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label lblyear 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
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
         Left            =   3240
         TabIndex        =   12
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   270
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label lblcourse 
         BackStyle       =   0  'Transparent
         Caption         =   "Surname, Firsname"
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
         Left            =   360
         TabIndex        =   10
         Top             =   3240
         Width           =   1935
      End
   End
   Begin VB.TextBox txtdate 
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.ListBox lstpos 
      Height          =   2985
      ItemData        =   "formresulta.frx":12F43
      Left            =   9120
      List            =   "formresulta.frx":12F62
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtcan 
      DataField       =   "StudentNo"
      DataSource      =   "adocandidate"
      Height          =   285
      Left            =   10440
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox txtstud 
      DataField       =   "StudentNo"
      DataSource      =   "adostud"
      Height          =   285
      Left            =   9360
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin MSAdodcLib.Adodc adocandidate 
      Height          =   375
      Left            =   9240
      Top             =   5160
      Width           =   4335
      _ExtentX        =   7646
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
      Connect         =   $"formresulta.frx":12FA7
      OLEDBString     =   $"formresulta.frx":13041
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM candidatesInfo"
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
   Begin MSAdodcLib.Adodc adostud 
      Height          =   375
      Left            =   9240
      Top             =   4680
      Width           =   4335
      _ExtentX        =   7646
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
      Connect         =   $"formresulta.frx":130DB
      OLEDBString     =   $"formresulta.frx":13175
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM personal_info"
      Caption         =   "adostud"
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
Attribute VB_Name = "formapply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub viewstud()
      With Me
       ' .cbopos.Clear
        .cbostudno.Clear
        .adostud.Refresh
        With .adostud.Recordset
          Do Until .EOF
            '.Filter = 0
           ' Do Until Me.adocandidate.Recordset.EOF
                Me.adocandidate.Recordset.Filter = 0
                Me.adocandidate.Recordset.Filter = "StudentNo LIKE '" & .Fields("StudentNo") & "'"
                If Me.adocandidate.Recordset.RecordCount = 0 Then
                    Me.cbostudno.AddItem .Fields("StudentNo")
              '  Else
               '     .MoveNext
                End If
                'If Me.adocandidate.Recordset.EOF Then
                   ' Exit Do
               ' End If
           ' Loop
          
            If .EOF Then
                Exit Do
                Exit Sub
            End If
            .MoveNext
            '.MoveNext
          Loop
        End With
    End With
End Sub

Private Sub Label7_Click()

End Sub

Private Sub cbopos_Change()
Me.lstpos.ListIndex = Me.cbopos.ListIndex
End Sub

Private Sub cbopos_Click()
Me.lstpos.ListIndex = Me.cbopos.ListIndex
End Sub

Private Sub cbostudno_Change()
With Me.adostud.Recordset
    .Filter = 0
    .Filter = "StudentNo LIKE '" & Me.cbostudno.Text & "'"
    Me.lblname.Caption = .Fields("Surname") & " " & .Fields("Firstname")
    Me.lblcourse.Caption = .Fields("Course")
    Me.lblyear.Caption = .Fields("year")
End With
End Sub

Private Sub cbostudno_Click()
With Me.adostud.Recordset
    .Filter = 0
    .Filter = "StudentNo LIKE '" & Me.cbostudno.Text & "'"
    Me.lblname.Caption = .Fields("Surname") & " " & .Fields("Firstname")
    Me.lblcourse.Caption = .Fields("Course")
    Me.lblyear.Caption = .Fields("year")
End With

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdsubmit_Click()
Dim askmeif  As Single
If Me.cbostudno.Text = "" Then
    MsgBox "To add new candidate choose his Student No.!!!", vbOKOnly + vbExclamation, "SSITE"
        Me.cbostudno.SetFocus
        Exit Sub
End If
If Me.cbopos.Text = "" Then
    MsgBox "To add new candidate choose his Position!!!", vbOKOnly + vbExclamation, "SSITE"
        Me.cbopos.SetFocus
        Exit Sub
End If
With Me.adocandidate.Recordset
    .AddNew
    .Fields("StudentNo") = Me.cbostudno.Text
    .Fields("PositionCode") = Me.lstpos.Text
    .Fields("DateField") = Me.txtdate.Text
    .Fields("SchoolYear") = MDIForm1.lblyear.Caption
    .Update
End With
Me.adocandidate.Refresh
MsgBox "Successfully adding new candidate!!!", vbOKOnly + vbInformation, "SSITE"
askmeif = MsgBox("Do you want to add new candidate again?", vbYesNo + vbQuestion, "SSITE")
    If askmeif = vbYes Then
        Call viewstud
        Me.Refresh
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
Dim i As Integer
Me.Move 0, 0
Me.txtdate.Text = Format$(Date, "MM/DD/YYYY")
Me.DTPicker1.Value = Format$(Date, "MM/DD/YYYY")
  Call viewstud
End Sub

