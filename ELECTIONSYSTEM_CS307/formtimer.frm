VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formtimer 
   Caption         =   "Election Status"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MouseIcon       =   "formtimer.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   11010
   ScaleWidth      =   12645
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   6960
      Width           =   8655
      Begin VB.Label Label10 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Course - Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time elapsed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   22
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "formtimer.frx":030A
      ScaleHeight     =   825
      ScaleWidth      =   12225
      TabIndex        =   20
      Top             =   0
      Width           =   12255
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
      Left            =   6840
      Picture         =   "formtimer.frx":B006
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9360
      Width           =   2535
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
      Height          =   735
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9360
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4080
      Width           =   8655
      Begin VB.Label Label8 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time elapsed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   16
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Course - Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   8655
      Begin VB.Label l1 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label l2 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Course - Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label l3 
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time elapsed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   0
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lvhigh 
      Height          =   2295
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
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
   End
   Begin VB.TextBox Text1 
      DataField       =   "StudentNo"
      DataSource      =   "adotime"
      Height          =   375
      Left            =   11040
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtstud 
      DataField       =   "StudentNo"
      DataSource      =   "adostud"
      Height          =   375
      Left            =   11040
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adostud 
      Height          =   375
      Left            =   12120
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   $"formtimer.frx":BCD0
      OLEDBString     =   $"formtimer.frx":BD6A
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Personal_info"
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
   Begin MSAdodcLib.Adodc adotime 
      Height          =   375
      Left            =   12120
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   $"formtimer.frx":BE04
      OLEDBString     =   $"formtimer.frx":BE9E
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
   Begin MSComctlLib.ListView lvslow 
      Height          =   2295
      Left            =   480
      TabIndex        =   17
      Top             =   4080
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
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
   End
   Begin MSComctlLib.ListView lvall 
      Height          =   2295
      Left            =   480
      TabIndex        =   25
      Top             =   6960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
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
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Student who voted and the time they consumed:"
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
      Left            =   480
      TabIndex        =   26
      Top             =   6600
      Width           =   5100
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Time of Voting per Student"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9240
      TabIndex        =   5
      Top             =   3480
      Width           =   2265
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblres 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   9240
      TabIndex        =   4
      Top             =   3960
      Width           =   2145
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Slowest Student who voted:"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   2970
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fastest Student who voted:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2910
   End
   Begin VB.Label lbldate 
      Caption         =   "date today"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Today is:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "formtimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timehigh, timelow, avetimevar, avetotal As Double
Dim topQ, tiectr As Integer
Dim target As String
Dim itm
Dim tsec, tmin, thour As Integer
Dim tres As String

Public Sub ConverTer()
    '--hour
    thour = topQ \ 3600
    topQ = topQ - (3600 * thour)
    '--min
    tmin = topQ \ 60
    topQ = topQ - (60 * tmin)
    '--sec
    tsec = topQ
    tres = Format$(thour, "00") & " : " & Format$(tmin, "00") & " : " & Format$(tsec, "00")
End Sub
Public Sub Timeall()
topQ = 0
   Me.lvall.ListItems.Clear
    With Me.adotime.Recordset
        Me.adotime.Refresh
        .Filter = "Datevote LIKE '" & Me.lbldate.Caption & "'"
        If .RecordCount <> 0 Then
       .MoveFirst
       Do Until .EOF
                target = .Fields("StudentNo")
                topQ = .Fields("timeconsume").Value
                With Me.adostud.Recordset
                    .Filter = "StudentNo LIKe '" & target & "'"
                        
                      Call ConverTer
                    '---------------------
                        With Me
                        Set itm = .lvall.ListItems.Add(, , Me.adostud.Recordset.Fields("Surname") & " , " & Me.adostud.Recordset.Fields("Firstname"))
                        itm.SubItems(1) = Me.adostud.Recordset.Fields("Course") & " - " & Me.adostud.Recordset.Fields("Year")
                        itm.SubItems(2) = tres
                        End With
                    '-----------------
                   .Filter = adFilterNone
                  End With
                .MoveNext
        Loop
         Me.adotime.Refresh
            Me.Refresh
           
    End If
    End With
 
End Sub
Public Sub infotimehigh()
tres = ""
tmin = 0
thour = 0
tsec = 0
topQ = 0
Me.lvhigh.ListItems.Clear
    With Me.adotime.Recordset
        Me.adotime.Refresh
        .Filter = "Datevote LIKE '" & Me.lbldate.Caption & "'"
        If .RecordCount <> 0 Then
       .MoveFirst
       .Sort = "Timeconsume ASC"
       target = .Fields("StudentNo")
        topQ = .Fields("timeconsume").Value
        Do Until .EOF
            If topQ = .Fields("TimeConsume").Value Then
                target = .Fields("StudentNo")
                With Me.adostud.Recordset
                    .Filter = "StudentNo LIKe '" & target & "'"
                     Call ConverTer
                    '---------------------
                        With Me
                        Set itm = .lvhigh.ListItems.Add(, , Me.adostud.Recordset.Fields("Surname") & " , " & Me.adostud.Recordset.Fields("Firstname"))
                        itm.SubItems(1) = Me.adostud.Recordset.Fields("Course") & " - " & Me.adostud.Recordset.Fields("Year")
                        itm.SubItems(2) = tres
                        End With
                    '-----------------
                    .Filter = adFilterNone
                  End With
            End If
          .MoveNext
        Loop
            Me.Refresh
            Me.adotime.Refresh
    End If
    End With
End Sub

Public Sub infotimelow()
topQ = 0
Me.lvslow.ListItems.Clear
    With Me.adotime.Recordset
        Me.adotime.Refresh
        .Filter = "Datevote LIKE '" & Me.lbldate.Caption & "'"
        If .RecordCount <> 0 Then
       .MoveFirst
       .Sort = "Timeconsume DESC"
       target = .Fields("StudentNo")
        topQ = .Fields("timeconsume").Value
                Do Until .EOF
                  If topQ = .Fields("TimeConsume").Value Then
                      target = .Fields("StudentNo")
                      With Me.adostud.Recordset
                          .Filter = "StudentNo LIKe '" & target & "'"
                          Call ConverTer
                          '---------------------
                              With Me
                              Set itm = .lvslow.ListItems.Add(, , Me.adostud.Recordset.Fields("Surname") & " , " & Me.adostud.Recordset.Fields("Firstname"))
                              itm.SubItems(1) = Me.adostud.Recordset.Fields("Course") & " - " & Me.adostud.Recordset.Fields("Year")
                              itm.SubItems(2) = tres
                              End With
                          '-----------------
                          .Filter = adFilterNone
                        End With
                  End If
                .MoveNext
        Loop
            Me.Refresh
            Me.adotime.Refresh
        End If
    End With
End Sub
Public Sub averagetime()
Me.adotime.Refresh
avetimevar = 0
With Me.adotime.Recordset
    .MoveFirst
    .Filter = "DateVote LIKE '" & Me.lbldate.Caption & "'"
    If .RecordCount <> 0 Then
        Do Until .EOF
            avetimevar = avetimevar + .Fields("timeConsume").Value
            topQ = (avetimevar / .RecordCount)
            .MoveNext
        Loop
            Me.adotime.Refresh
    End If
End With
    Call ConverTer
    Me.lblres.Caption = tres
End Sub


Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdexit_Click()
Call exitQ
End Sub

Private Sub Form_Activate()
Me.lbldate.Caption = Format(Date, "MM/DD/YYYY")
Call infotimehigh
Call infotimelow
Call averagetime
Call Timeall
End Sub

Private Sub Form_Load()
With Me.lvhigh
    .ColumnHeaders(1).Width = Me.l1.Width
    .ColumnHeaders(2).Width = Me.l2.Width
    .ColumnHeaders(3).Width = Me.l3.Width
End With
With Me.lvslow
    .ColumnHeaders(1).Width = Me.l1.Width
    .ColumnHeaders(2).Width = Me.l2.Width
    .ColumnHeaders(3).Width = Me.l3.Width
End With
With Me.lvall
    .ColumnHeaders(1).Width = Me.l1.Width
    .ColumnHeaders(2).Width = Me.l2.Width
    .ColumnHeaders(3).Width = Me.l3.Width
End With
End Sub

