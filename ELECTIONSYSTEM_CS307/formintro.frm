VERSION 5.00
Begin VB.Form formintro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "formintro.frx":0000
   ScaleHeight     =   2115
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4800
      Top             =   1560
   End
End
Attribute VB_Name = "formintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Form1.Show
Unload Me
End Sub
