VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   ControlBox      =   0   'False
   DrawMode        =   4  'Mask Not Pen
   DrawStyle       =   6  'Inside Solid
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   405
      Left            =   11880
      TabIndex        =   15
      Top             =   3960
      Width           =   1350
   End
   Begin VB.CommandButton Cmd_Simpan 
      Caption         =   "Simpan"
      Height          =   405
      Left            =   13560
      TabIndex        =   14
      Top             =   4560
      Width           =   1350
   End
   Begin VB.CommandButton Cmd_Batal 
      Caption         =   "Batal"
      Height          =   405
      Left            =   13560
      TabIndex        =   13
      Top             =   3960
      Width           =   1350
   End
   Begin VB.CommandButton cmd_hapus 
      Caption         =   "hapus"
      Height          =   405
      Left            =   11880
      TabIndex        =   12
      Top             =   4560
      Width           =   1350
   End
   Begin VB.TextBox tnormal 
      Height          =   330
      Left            =   11880
      TabIndex        =   5
      Top             =   3240
      Width           =   3045
   End
   Begin VB.TextBox trusak 
      Height          =   330
      Left            =   11880
      TabIndex        =   4
      Top             =   2760
      Width           =   3045
   End
   Begin VB.TextBox tjumlah 
      Height          =   330
      Left            =   11880
      TabIndex        =   3
      Top             =   2280
      Width           =   3045
   End
   Begin VB.TextBox tharga 
      Height          =   330
      Left            =   11880
      TabIndex        =   2
      Top             =   1800
      Width           =   3045
   End
   Begin VB.TextBox tnama 
      Height          =   330
      Left            =   11880
      TabIndex        =   1
      Top             =   1320
      Width           =   3045
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   65535
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   22
      TabAction       =   2
      RowDividerStyle =   5
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "no"
         Caption         =   "No."
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
         DataField       =   "nama"
         Caption         =   "Nama Barang"
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
      BeginProperty Column02 
         DataField       =   "harga"
         Caption         =   "Harga"
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
      BeginProperty Column03 
         DataField       =   "jmltotal"
         Caption         =   "Jumlah Total"
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
      BeginProperty Column04 
         DataField       =   "rusak"
         Caption         =   "Rusak"
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
      BeginProperty Column05 
         DataField       =   "normal"
         Caption         =   "Normal"
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
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1484.787
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Inventaris Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   11
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga  :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   10
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   9
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rusak  :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   8
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total  :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   7
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama  :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10800
      TabIndex        =   6
      Top             =   1320
      Width           =   1260
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rcbarang As Recordset
Attribute rcbarang.VB_VarHelpID = -1
Dim rcjml As Recordset

Private Sub Cmd_Batal_Click()
Set DataGrid1.DataSource = Nothing
rcbarang.CancelUpdate
Set DataGrid1.DataSource = rcbarang
tjumlah = ""
trusak = ""
tnormal = ""
tnama = ""
tharga = ""

End Sub

Private Sub cmd_hapus_Click()
tanya = MsgBox("Anda yakin untuk menghapus data ini ?", vbYesNo + vbQuestion, "Konfirmasi")
If tanya = vbNo Then Exit Sub
If rcbarang.RecordCount = 0 Then Exit Sub

rcbarang.Delete adAffectCurrent
rcbarang.Update
End Sub

Private Sub Cmd_Simpan_Click()
If tnama = "" Or tharga = "" Or tjumlah = "" Then
MsgBox "Data Tidak Boleh Kosong"
Exit Sub
End If
tanya = MsgBox("Simpan Data ?", vbYesNo + vbQuestion, "Konfirmasi")
If tanya = vbNo Then Exit Sub

rcbarang!nama = tnama
If tnama = "" Then Exit Sub
rcbarang!harga = tharga
rcbarang!nama = tnama
rcbarang!jmltotal = tjumlah
rcbarang!rusak = trusak
rcbarang!Normal = tnormal
rcbarang.Update
End Sub

Private Sub Command1_Click()
rcbarang.AddNew
End Sub

Private Sub Form_Load()
Call Aktif_Koneksi
Set rcbarang = New Recordset
rcbarang.Open "select * from daftarbarang", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rcbarang
End Sub

Private Sub Label7_Click()

End Sub

Private Sub rcbarang_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If rcbarang.RecordCount > 0 Then
tnama = rcbarang!nama & ""
tharga = rcbarang!harga & ""
tjumlah = rcbarang!jmltotal & ""
trusak = rcbarang!rusak & ""
tnormal = rcbarang!Normal & ""
End If
End Sub

