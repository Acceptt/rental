VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Cetak "
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   4920
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   11055
      _ExtentX        =   19500
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
         DataField       =   "periode"
         Caption         =   "Periode"
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
         DataField       =   "penjualan"
         Caption         =   "Penjualan"
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
         DataField       =   "nilaix"
         Caption         =   "X"
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
         DataField       =   "xy"
         Caption         =   "XY"
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
         DataField       =   "xx"
         Caption         =   "XX"
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
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1890.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1934.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1679.811
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Client Wisma Rias"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label7 
      Caption         =   "Label1"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Label1"
      Height          =   495
      Left            =   10920
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents rcramal As Recordset
Attribute rcramal.VB_VarHelpID = -1
Dim rcjml As Recordset


Private Sub Command1_Click()
Label6.Caption = (Label1.Caption / Label2.Caption) + ((Label4.Caption / Label5.Caption) * (Label7.Caption + 1))

End Sub

Private Sub Command2_Click()
Dim rccetak As New Recordset
rccetak.Open "select * from peramalan", db, adOpenDynamic, adLockOptimistic

With frm_report
.VSReport1.Load App.Path & "\report.xml", "pinjaman"
.VSReport1.DataSource.Recordset = rccetak
.VSReport1.Render .VSPrinter1
.Show
End With
End Sub

Private Sub Form_Load()
Call Aktif_Koneksi
Set rcramal = New Recordset
rcramal.Open "select * from peramalan", db, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rcramal
If rcramal.RecordCount < 1 Then
MsgBox "data kosong"
Exit Sub
End If
  rcramal.MoveFirst
   Do Until rcramal.EOF
            jumlah = jumlah + rcramal!penjualan
            x = 0 + rcramal!nilaix
            xy = xy + rcramal!xy
            xx = xx + rcramal!xx
            xxx = 0 + rcramal!nilaix
        rcramal.MoveNext '
            Label1.Caption = jumlah
            Label3.Caption = x
            Label4.Caption = xy
            Label5.Caption = xx
            Label7.Caption = xxx
    Loop
jumlah1 = rcramal.RecordCount
Label2.Caption = jumlah1


End Sub
