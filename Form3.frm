VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   120
   ClientWidth     =   14115
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   14115
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   4080
      TabIndex        =   5
      Top             =   4320
      Width           =   7335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   4080
      TabIndex        =   4
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   7335
      Begin VB.Label Label3 
         Caption         =   "2. Kursi"
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "1. Gedung"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paket Budged"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paket Perumahan"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paket Gedung"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilihan Paket"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub
