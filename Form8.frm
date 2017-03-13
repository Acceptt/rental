VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10815
   LinkTopic       =   "Form8"
   ScaleHeight     =   5040
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   65
      Left            =   10920
      Top             =   1440
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    ProgressBar1.Value = ProgressBar1.Value + 2
    Label1.Caption = "Loading Complite . . ."
    Label2.Caption = ProgressBar1.Value & "%"
    If (ProgressBar1.Value = ProgressBar1.Max) Then

       Timer1.Enabled = False
       Unload Me
       login.Show
    End If
End Sub

