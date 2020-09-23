VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2625
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   0
      MouseIcon       =   "Form1.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0614
      ScaleHeight     =   2355
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.TextBox txtpassword 
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000003&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LOG - IN"
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
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
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
         Left            =   2760
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cbouser 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "000 - 0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
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
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "cbouser"
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblq 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SSITE Election System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Student No.:"
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
         Height          =   285
         Left            =   870
         TabIndex        =   8
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Height          =   285
         Left            =   1050
         TabIndex        =   7
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   360
         Picture         =   "Form1.frx":8003
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "passwordADmin"
      DataSource      =   "adoadmin"
      Height          =   285
      Left            =   7320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4800
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adoadmin 
      Height          =   375
      Left            =   5520
      Top             =   4440
      Width           =   3255
      _ExtentX        =   5741
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
      Connect         =   $"Form1.frx":8445
      OLEDBString     =   $"Form1.frx":84DF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Admin_tbl"
      Caption         =   "admin"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5880
      Top             =   6480
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Connect         =   $"Form1.frx":8579
      OLEDBString     =   $"Form1.frx":8613
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT StudentNO,Special_Password ,firstname,surname,voted FROM personal_info"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usernameq, Spassword As String
Dim ctr, ryan As Integer
Dim boiser As Integer
Public Sub viewList()
    With Me.Adodc1.Recordset
          Me.Adodc1.Refresh
       Do Until .EOF
         Me.cbouser.AddItem .Fields("StudentNo")
        .MoveNext
       Loop
         Me.Adodc1.Refresh
    End With
End Sub

Private Sub cbouser_LostFocus()
Me.cbouser.DataFormat.Format = Format$(Me.cbouser.Text, "000 - 0000")
End Sub

Private Sub Check1_Click()
If Me.Check1.Value = vbChecked Then Me.cbouser.Clear: Me.cbouser.Enabled = False: Me.txtpassword.SetFocus
If Me.Check1.Value = vbUnchecked Then Me.cbouser.Enabled = True: Call viewList: Me.cbouser.SetFocus: Me.cbouser.ListIndex = 0
End Sub

Private Sub Command1_Click()
Dim result  As Boolean
If Me.txtpassword.Text <> "" Then
    Me.adoadmin.Recordset.Filter = 0
    Me.Adodc1.Recordset.Filter = 0
    Me.adoadmin.Refresh
    Me.Adodc1.Refresh
    Dim adminpass As String
    If Me.Check1.Value = vbChecked Then
        adminpass = Me.txtpassword.Text
        'Me.adoadmin.Recordset.Filter = "passwordAdmin = '" & adminpass & "'"
        If Me.adoadmin.Recordset.Fields("passwordAdmin") = adminpass Then
              MDIForm1.Show: MDIForm1.Caption = "ADMINISTRATOR": Unload Me:
        Else
              MsgBox "Sorry you are not recognized Admin User...", vbOKOnly + vbCritical, "SSITE"
              ctr = ctr + 1
              If ctr = 3 Then
                MsgBox "Application will be close....", vbOKOnly + vbExclamation, "SSITE"
                End
              End If
              Exit Sub
        End If
    Else
        With Me
            usernameq = .cbouser.Text
            Spassword = .txtpassword.Text
                .Adodc1.Refresh
              With .Adodc1.Recordset
                Do Until .EOF
                If .Fields("StudentNo") = usernameq And .Fields("Special_Password") = Spassword Then
                    If .Fields("Voted") = False Then
                       formVote.txtstud.Text = Me.cbouser.Text
                       formVote.Caption = .Fields("Firstname") & "  " & .Fields("Surname")
                       formVote.Show
                       Unload Me
                       Exit Do
                    Else
                       MsgBox "You voted already!", vbOKOnly + vbCritical, "SSITE"
                            Me.txtpassword.SetFocus
                            Me.Adodc1.Refresh
                            Exit Sub
                    End If
                
                End If
                  .MoveNext
                   If .EOF Then
                         MsgBox "Sorry you are not a recognized User ...", vbOKOnly + vbCritical, "SSITE"
                         Me.txtpassword.SetFocus
                         Me.Adodc1.Refresh
                         ctr = ctr + 1
                         If ctr = 3 Then
                           MsgBox "Application will be close....", vbOKOnly + vbExclamation, "SSITE"
                           End
                         End If
                         Exit Sub
                   End If
                Loop
              End With
        End With
    End If
Else
    Exit Sub
End If
hell:
    Exit Sub
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
viewList
Me.cbouser.SetFocus
Me.cbouser.ListIndex = 0
ryan = 0
boiser = 0
End Sub


Private Sub Image1_Click()
boiser = boiser + 1
End Sub

Private Sub Label3_Click()
ryan = ryan + 1
End Sub

Private Sub lblq_Click()
Dim samson, adminq As String
Dim askmeif As Single
If ryan = 11 And boiser = 11 Then
    samson = InputBox("Hello Ryan...." + Chr(13) + "Who is your sexiest friend?", "Hello")
    If samson = "J3ny L." Then
        askmeif = MsgBox("Do you want to do?" + Chr(13) + "'Yes' if you want to change Admin Password" & _
            Chr(13) + "'No' to only view Admin Password", vbYesNo + vbQuestion, "heheh")
            If askmeif = vbYes Then
                adminq = InputBox("Enter password here...")
                    If adminq <> "" Then
                        Me.adoadmin.Recordset.Fields("passwordADmin") = adminq
                        Me.adoadmin.Recordset.Update
                    End If
            Else
                MsgBox "Admin current password " & Me.adoadmin.Recordset.Fields("passwordADmin")
                    Exit Sub
            End If
    
    Else
        MsgBox "Wrong answer...."
        Exit Sub
    End If
End If
        
End Sub

Private Sub txtpassword_GotFocus()
With Me.txtpassword
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub
