VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Log In"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tcari 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox pas 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmd_login 
      Caption         =   "Sign in"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmd_next 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox tpas 
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox tus 
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox ttingkat 
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox tsr 
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2778
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selamat Datang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   1080
      TabIndex        =   11
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Silahkan Masukkan Username anda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rcpass As Recordset
Attribute rcpass.VB_VarHelpID = -1

Private Sub cmd_login_Click()
If pas = tpas And ttingkat = "admin" Then
MDIForm1.Show
'ElseIf pas = tpas And ttingkat = "user" Then
'Form11.Show
Else
MsgBox "Password tidak cucok"
Exit Sub
End If

Unload Me
End Sub

Private Sub cmd_next_Click()
tus = ""

Call Aktif_Koneksi
Set rcpass = New Recordset
Dim sql As String
sql = "select * from admin where user like '" & Tcari.Text & "'"
rcpass.Open sql, db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rcpass
If tus = "" Then
MsgBox "User Tidak dikenali"
Exit Sub
End If
pas.Visible = True
cmd_login.Visible = True
cmd_next.Visible = False
cmd_next.Default = False
cmd_login.Default = True
Tcari.Enabled = False
label.Caption = Tcari
Tcari.Visible = False


End Sub


Private Sub Command1_Click()
If Text2 = "" Then Exit Sub
If Text4 = Text1 Then
Tcari.Visible = True
cmd_next.Visible = True
Command1.Visible = False
Text2.Visible = False
Tcari.Enabled = True
cmd_next.Enabled = True
Else
MsgBox "Silahkan Coba Kembali"
Tcari.Enabled = False
cmd_next.Enabled = False

End If
rcpass!sr = Text4
rcpass.Update

End Sub

Private Sub Form_Load()
Aktif_Koneksi
Set rcpass = New Recordset
rcpass.Open "select * from admin", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rcpass

End Sub



Private Sub rcpass_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If rcpass.RecordCount > 0 Then
ttingkat = rcpass!Level & ""
tus = rcpass!user & ""
tpas = rcpass!Password & ""
'tsr = rcpass!sr & ""
End If
End Sub





