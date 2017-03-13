VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13815
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   13815
   WindowState     =   2  'Maximized
   Begin VB.TextBox tpen2 
      Height          =   375
      Left            =   9120
      TabIndex        =   38
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox tper2 
      Height          =   375
      Left            =   9120
      TabIndex        =   37
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox tx2 
      Height          =   375
      Left            =   9120
      TabIndex        =   36
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txy2 
      Height          =   375
      Left            =   9120
      TabIndex        =   35
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txx2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   9120
      TabIndex        =   34
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txx 
      Height          =   375
      Left            =   6960
      TabIndex        =   33
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txy 
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox tx 
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      TabIndex        =   28
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox tx1 
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox tper1 
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox tpen1 
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox tpen 
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox tper 
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox ttotal 
      Height          =   615
      Left            =   5520
      TabIndex        =   21
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox tperiode 
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox ttanggal 
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   2760
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1215
      Left            =   5640
      TabIndex        =   18
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2143
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
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   4440
      List            =   "Form4.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":0035
      Left            =   2880
      List            =   "Form4.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox tlama 
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox tsewa 
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox thp 
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox talamat 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox tnama 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   3255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Paket Budget"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Paket Perumahan"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3960
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Paket Gedung"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1215
      Left            =   5640
      TabIndex        =   22
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2143
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1215
      Left            =   8400
      TabIndex        =   30
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
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
   Begin VB.Label jml 
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lama Sewa "
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Sewa"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No HP"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Paket"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rcclient As Recordset
Dim WithEvents rcramal As Recordset
Attribute rcramal.VB_VarHelpID = -1
Dim WithEvents rcdata As Recordset
Attribute rcdata.VB_VarHelpID = -1

Private Sub Combo1_Click()
If tnama = "" Or talamat = "" Or tsewa = "" Then
MsgBox "Data Tidak Boleh Kosong"
Exit Sub
Combo2.Clear
End If

End Sub

Private Sub Combo2_Click()
ttanggal = tsewa + Combo1 + Combo2
End Sub

Private Sub Command1_Click()
If tnama = "" Or talamat = "" Or tsewa = "" Then
MsgBox "Data Tidak Boleh Kosong"
Exit Sub
End If
tanya = MsgBox("Simpan Data ?", vbYesNo + vbQuestion, "Konfirmasi")
If tanya = vbNo Then Exit Sub
rcclient.AddNew
rcclient!nama = tnama
rcclient!alamat = talamat
rcclient!tanggal = ttanggal
rcclient!nominal = ttotal
rcclient.Update
If tper = "" Then
rcramal.AddNew
rcramal!periode = tper1
rcramal!penjualan = tpen1
rcramal!nilaix = tx
rcramal!xy = tx * tpen1
rcramal!xx = tx * tx
rcramal.Update
Else
rcramal!periode = tper
rcramal!penjualan = tpen1
rcramal.Update
End If

tnama = ""
talamat = ""
tsewa = ""
thp = ""
tperiode = ""
ttanggal = ""
tpen = ""
tpen1 = ""
tper = ""
tper1 = ""
End Sub



Private Sub Form_Load()
Aktif_Koneksi
Set rcclient = New Recordset
rcclient.Open "select * from client", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rcclient


Set rcdata = New Recordset
rcdata.Open "select * from peramalan order by no desc", db, adOpenDynamic, adLockOptimistic
Set DataGrid2.DataSource = rcdata
'Dim jmlh As Integer
'jmlh = rcramal.RecordCount
'jml.Caption = jmlh


End Sub

Private Sub Option1_Click()
ttotsl = 5000
End Sub

Private Sub tcoba_Change()
ttotal = tcoba
End Sub

Private Sub rcdata_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If rcdata.RecordCount > 0 Then
tx2 = rcdata!nilaix & ""
End If
End Sub

Private Sub rcramal_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If rcramal.RecordCount > 0 Then
tper = rcramal!periode & ""
tpen = rcramal!penjualan & ""
tx1 = rcramal!nilaix & ""
txy2 = rcramal!xy & ""
txx2 = rcramal!xx & ""
Else
tper = ""
tpen = ""
tper1 = tperiode
tpen1 = ttotal
End If
End Sub

Private Sub thp_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        'MsgBox "Isikan Angka Saja", 48, "Perhatian"
        KeyAscii = 0
    End If
End Sub

Private Sub tlama_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        'MsgBox "Isikan Angka Saja", 48, "Perhatian"
        KeyAscii = 0
    End If
End Sub


Private Sub tperiode_Change()
If tperiode = "" Then
Exit Sub
End If
Call Aktif_Koneksi
Set rcramal = New Recordset
rcramal.Open "select * from peramalan where periode like '%" & tperiode.Text & "%' order by periode desc", db, adOpenDynamic, adLockOptimistic
Set DataGrid3.DataSource = rcramal
Dim jmlh As Integer
jmlh = rcramal.RecordCount
jml.Caption = jmlh
End Sub

Private Sub tsewa_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace) Then
        'MsgBox "Isikan Angka Saja", 48, "Perhatian"
        KeyAscii = 0
    End If
End Sub

Private Sub ttanggal_Change()
tperiode = Combo1 + Combo2
End Sub

Private Sub ttotal_Change()
If tx2 = "" Then
tx = 0
ElseIf tx2 = 0 And tper1 = "" Then
tx = tx1
ElseIf tx1 = "" Then
tx = 1
ElseIf tper1 = "" Then
tx = tx1
Else
Exit Sub
End If
If tpen = "" Then
tpen1 = ttotal
Else
    If ttotal = "" Then
        tpen1 = ""
    Else
        tpen1 = Val(tpen) + Val(ttotal)
    End If
End If
If tper = "" And tpen = "" Then
tx = (tx2) + 2
End If
End Sub

Private Sub tx_Change()
txx = tx * tx
End Sub

