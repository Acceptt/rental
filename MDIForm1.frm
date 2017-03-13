VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Aplikasi Jasa Rental"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10890
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   0
      ScaleHeight     =   2.25
      ScaleMode       =   4  'Character
      ScaleWidth      =   90.25
      TabIndex        =   8
      Top             =   5775
      Width           =   10890
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   10830
      TabIndex        =   0
      Top             =   0
      Width           =   10890
      Begin VB.CommandButton home 
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton transaksi 
         Caption         =   "Transaksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton db 
         Caption         =   "Data Barang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton dc 
         Caption         =   "Daftar Client"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton prediksi 
         Caption         =   "Prediksi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton logout 
         Caption         =   "Log Out"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10440
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton dp 
         Caption         =   "Daftar Paket"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Menu lapo 
      Caption         =   "laporan"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub db_Click()
Unload Form2
Unload Form3
Unload Form6
Unload Form4
Form5.Show
End Sub

Private Sub dc_Click()
Unload Form2
Unload Form6
Unload Form5
Unload Form4
Unload Form3
Form7.Show
End Sub

Private Sub dp_Click()
Unload Form2
Unload Form6
Unload Form5
Unload Form4
Unload Form7
Form3.Show
End Sub

Private Sub home_Click()
'Picture2
Form2.Show
Unload Form7
Unload Form6
Unload Form3
Unload Form5
Unload Form4
End Sub

Private Sub logout_Click()
tanya = MsgBox("Anda yakin untuk Keluar ?", vbYesNo + vbQuestion, "Konfirmasi")
If tanya = vbNo Then Exit Sub
Form1.Show
Unload Me
End Sub

Private Sub prediksi_Click()
Unload Form2
Unload Form3
Unload Form7
Unload Form5
Unload Form4
Form6.Show
End Sub

Private Sub transaksi_Click()
Unload Form2
Unload Form3
Unload Form5
Unload Form6
Unload Form7
Form4.Show
End Sub
