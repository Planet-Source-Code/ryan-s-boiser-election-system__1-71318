VERSION 5.00
Begin VB.Form formfind 
   BorderStyle     =   0  'None
   Caption         =   "Find..."
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "formfind.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "formfind.frx":030A
   ScaleHeight     =   2625
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Student No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancel 
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
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO!"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtfind 
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblformat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   60
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Format should be:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Find what?"
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
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "formfind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdgo_Click()
If Me.txtfind.Text = "" Then
   ' MsgBox "Finding what? a null?", vbOKOnly + vbCritical, "SSITE"
        Exit Sub
End If
Dim myCriteria, mytarget As String
If Me.Option1.Value = True Then myCriteria = "StudentNo"
If Me.Option2.Value = True Then myCriteria = "Surname"
mytarget = Me.txtfind.Text
'Form2.AdopersonalINfo.Refresh
With formstudent.AdopersonalINfo.Recordset
    .Find myCriteria & " LIKE '" & mytarget & "*'"
    If .EOF Then
        MsgBox "Sorry record not found...", vbOKOnly + vbSystemModal, "SSITE"
            Me.txtfind.Text = ""
            formstudent.AdopersonalINfo.Refresh
            Me.txtfind.SetFocus
            Exit Sub
    Else
        formstudent.Show
        Unload Me
    End If
End With
'Me.txtfind.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
        Me.Move 0, 0
        Me.lblformat.Caption = "0611544"
        Me.txtfind.SetFocus
        'Me.Top = (ScaleHeight - lvRent.Top)
        'Me.Top = 0
End Sub

Private Sub Option1_Click()
Me.lblformat.Caption = "0611544"
Me.txtfind.SetFocus
End Sub

Private Sub Option2_Click()
Me.lblformat.Caption = "Boiser"
Me.txtfind.SetFocus
End Sub

Private Sub txtfind_Change()
'cmdgo_Click
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgo_Click
End If
End Sub
