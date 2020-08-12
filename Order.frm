VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Order 
   Caption         =   "Data Pesanan Barang"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5280
      Top             =   6240
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   4800
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Order.frx":0000
      Height          =   2850
      Left            =   120
      TabIndex        =   37
      Top             =   3240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5027
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "NO"
         Caption         =   "NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "TANGGAL"
         Caption         =   "TANGGAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "KODE"
         Caption         =   "KODE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "KETERANGAN"
         Caption         =   "KETERANGAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "JUMLAH"
         Caption         =   "JUMLAH"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker TGLPO 
      Height          =   345
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92602369
      CurrentDate     =   39561
   End
   Begin VB.CommandButton CMDRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   36
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox DP 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6600
      TabIndex        =   6
      Top             =   840
      Width           =   1800
   End
   Begin VB.ComboBox CBOKodePms 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   3750
   End
   Begin VB.TextBox Alamat 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   3750
   End
   Begin VB.TextBox Telepon 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   3750
   End
   Begin VB.TextBox NamaPms 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   3750
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   6240
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      Left            =   8520
      TabIndex        =   9
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1800
      TabIndex        =   13
      Top             =   6240
      Width           =   750
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   960
      TabIndex        =   12
      Top             =   6240
      Width           =   750
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   750
   End
   Begin VB.TextBox Keterangan 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1440
      TabIndex        =   8
      Top             =   2760
      Width           =   5000
   End
   Begin VB.TextBox Jumlah 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6840
      TabIndex        =   10
      Top             =   2760
      Width           =   1540
   End
   Begin VB.TextBox JumlahPO 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6600
      TabIndex        =   5
      Top             =   480
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker TGLTran 
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   93192193
      CurrentDate     =   39561
   End
   Begin VB.Label NomorPO 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   38
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   35
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   34
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Biaya"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      TabIndex        =   33
      Top             =   6240
      Width           =   1245
   End
   Begin VB.Label Sisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   32
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Ket 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   31
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keterangan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   30
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa / Piutang"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   29
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " DP"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   28
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Laba"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   27
      Top             =   6240
      Width           =   1245
   End
   Begin VB.Label Laba 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   26
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   25
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Pemesan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Pemesan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tgl Transaksi"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label Biaya 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7080
      TabIndex        =   20
      Top             =   6600
      Width           =   1245
   End
   Begin VB.Label JmlData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   19
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uraian Transaksi"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   18
      Top             =   2400
      Width           =   4995
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6840
      TabIndex        =   17
      Top             =   2400
      Width           =   1545
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah "
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   16
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal PO"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5280
      TabIndex        =   15
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor PO"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1250
   End
End
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Me.Height = Me.Height + 100
Tengah
If Me.Height >= 8000 Then
    Timer1.Enabled = False
    Tengah
End If
End Sub

Public Sub Tengah()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Activate()
'judul form ada tambahan nama user dari login
Order.Caption = "Data Pesanan " & Login.TxtNamaKsr
'buka database
Call BukaDB
'hub adodc ke database
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBAKN.mdb"
'hub adodc ke tabel transaksi
Adodc1.RecordSource = "Transaksi"
Adodc1.Refresh
'hub datagrid ke adodc
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
'kosongkan dulu tabel transaksi
Call TBLKosong
Adodc1.Recordset.MoveFirst
'panggil nomor po otomatis
Call AutoPO
'panggil nomor pemesan otomatis
Call AutoPMS
End Sub

Sub Form_Load()
Me.Height = 10
Call BukaDB
TglPO = Format(Date, "dd-mm-yyyy")
TGLTran = Format(Date, "dd-mm-yyyy")
RSPerkiraan.Open "select * from perkiraan", Conn
List1.Clear
Do While Not RSPerkiraan.EOF
    'tampilkan kode dan nama perkiraan di list1
    List1.AddItem RSPerkiraan!KodePrk & Space(3) & RSPerkiraan!NamaPrk
    RSPerkiraan.MoveNext
Loop

Call BukaDB
RSPemesan.Open "select * from pemesan", Conn
CBOKodePms.Clear
Do While Not RSPemesan.EOF
    'tampilkan kode dan nama pemesan di
    CBOKodePms.AddItem RSPemesan!KodePms & Space(5) & RSPemesan!NamaPms
    RSPemesan.MoveNext
Loop
End Sub

'jumlah PO tidak boleh kosong
Private Sub JumlahPO_KeyPress(Keyascii As Integer)
On Error Resume Next
If Keyascii = 13 Then
    JumlahPO = Format(JumlahPO, "###,###,###")
    If JumlahPO = "" Or JumlahPO = 0 Then
        MsgBox "Jumlah harus diisi"
        JumlahPO.SetFocus
        Exit Sub
    Else
        DP.SetFocus
    End If
End If
'hanya dpt diisi angka
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

'masalah DP (uang muka)
Private Sub DP_KeyPress(Keyascii As Integer)
On Error Resume Next
If Keyascii = 13 Then
    'jika dp 0 atau kosong maka
    If DP = 0 Or DP = "" Then
        DP = 0
        Sisa = JumlahPO
        Kembali = 0
        Ket = "BELUM LUNAS"
        TGLTran.SetFocus
    Else
        'jika dp lebih kecil dari jumlah PO, maka
        If Val(DP) < JumlahPO Then
            'tampilkan sisanya
            Sisa = Format(JumlahPO - DP, "###,###,###")
            Kembali = 0
            Ket = "BELUM LUNAS"
            TGLTran.SetFocus
            DP = Format(DP, "###,###,###")
        'jika dp = jumlahpo maka...
        ElseIf Val(DP) = JumlahPO Then
            Sisa = 0
            Kembali = 0
            Ket = "LUNAS"
            TGLTran.SetFocus
            DP = Format(DP, "###,###,###")
        'jika DP lebih besar dari jumlahPo maka...
        ElseIf Val(DP) > JumlahPO Then
            Sisa = 0
            Kembali = Format(DP - JumlahPO, "###,###,###")
            Ket = "LUNAS"
            TGLTran.SetFocus
            DP = Format(DP, "###,###,###")
        End If
    End If
End If
End Sub

Private Sub CMDRefresh_Click()
Call AutoPMS
Call BersihkanPms
End Sub

'kode pemesan tidak boleh ksoong
Private Sub cboKodePms_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If CBOKodePms = "" Then
        MsgBox "kode pemesan kosong"
        CBOKodePms.SetFocus
        Exit Sub
    Else
        Call BukaDB
        'buka tabel pemesan yang kodenya dipilih dalam combo (6 digit dari kiri)
        RSPemesan.Open "select * from pemesan where kodepms='" & Left(CBOKodePms, 6) & "'", Conn
        If Not RSPemesan.EOF Then
            Call MatiPms
            Call DataPms
            TglPO.SetFocus
        Else
            BersiTelp3ms
            HidupPMS
            NamaPms.SetFocus
        End If
    End If
End If
End Sub

Private Sub cbokodepms_Click()
Call BukaDB
RSPemesan.Open "select * from pemesan where kodepms='" & Left(CBOKodePms, 6) & "'", Conn
If Not RSPemesan.EOF Then
    Call MatiPms
    Call DataPms
Else
    MsgBox "kode pemesan tidak terfadtar"
    TglPO.SetFocus
End If
End Sub

Private Sub NomorPO_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSMasterPO.Open "select * from masterpo where NOPO='" & NomorPO & "'", Conn
    If Not RSMasterPO.EOF Then
        TglPO = RSMasterPO!TglPO
        Jumlah = RSMasterPO!Jumlah
        DP = RSMasterPO!DP
        Sisa = RSMasterPO!Sisa
        Telp3P = RSMasterPO!Telp3P
        CBOKodePms = RSMasterPO!KodePms
        Ket = RSMasterPO!Ket
    Else
        TglPO.SetFocus
    End If
End If
End Sub


Private Sub AutoPO()
Call BukaDB
RSMasterPO.Open "select * from MasterPO Where NOPO In(Select Max(NOPO)From MasterPO)Order By NOPO Desc", Conn
RSMasterPO.Requery
    Dim Urutan As String * 11
    Dim Hitung As Long
    With RSMasterPO
        If .EOF Then
            Urutan = "PO-" + Format(Date, "yymmdd") + "01"
            NomorPO = Urutan
        Else
            If Mid(!NOPO, 4, 6) <> Format(Date, "yymmdd") Then
                Urutan = "PO-" + Format(Date, "yymmdd") + "01"
            Else
                Hitung = Right(!NOPO, 8) + 1
                Urutan = "PO-" + Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        NomorPO = Urutan
    End With
End Sub

'fungsi untuk mencari kode pemesan otomatis
Private Sub AutoPMS()
Call BukaDB
RSPemesan.Open ("select * from pemesan Where KodePms In(Select Max(KodePms)From pemesan)Order By KodePms Desc"), Conn
RSPemesan.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RSPemesan
        If .EOF Then
            Urutan = "PMS" + "001"
            CBOKodePms = Urutan
        Else
            Hitung = Right(!KodePms, 3) + 1
            Urutan = "PMS" + Right("000" & Hitung, 3)
        End If
        CBOKodePms = Urutan
    End With
End Sub

Sub HidupPMS()
NamaPms.Enabled = True
Alamat.Enabled = True
Telepon.Enabled = True
End Sub

Private Sub NamaPms_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If NamaPms = "" Then
        MsgBox "nama pemesan kosong"
        NamaPms.SetFocus
        Exit Sub
    Else
        Alamat.SetFocus
    End If
End If
End Sub

Private Sub alamat_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Alamat1 = "" Then
        Alamat1 = "-"
        Telepon.SetFocus
    Else
        Telepon.SetFocus
    End If
End If
End Sub


Private Sub Telepon_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Telepon = "" Then
        Telepon = "-"
        TglPO.SetFocus
    Else
        TglPO.SetFocus
    End If
End If
End Sub

Private Sub TglPO_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then JumlahPO.SetFocus
End Sub

Private Sub Keterangan_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Keterangan = "" Then
        MsgBox "Uraian transaksi masih kosong"
        Keterangan.SetFocus
        Exit Sub
    Else
        List1.SetFocus
    End If
End If
End Sub

'fungsi ini digunakan untuk mengedit jumlah dalam transaksi
'atau dalam grid jika laba lebih kecil dari biaya
Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
If DataGrid1.Col = 4 Then
    Call TotalBiaya
    Laba = Format(JumlahPO - Biaya, "#,###,###")
    DataGrid1.Refresh
End If
End Sub

Private Sub CmdSimpan_Click()
'On Error Resume Next
Pesan = MsgBox("Data sudah benar", vbYesNo)
If Pesan = vbYes Then
    'cegah data kosong
    If JumlahPO = "" Or DP = "" Or Sisa = "" Or Ket = "" Or CBOKodePms = "" Or Biaya = "" Or Laba = "" Then
        MsgBox "Data belum lengkap, cek kembali dengan teliti"
        Exit Sub
    End If
    
    'simpan ke tabel pemesan jika kodenya baru
    Call BukaDB
    RSPemesan.Open "select * from Pemesan where kodePMS='" & Left(CBOKodePms, 6) & "'", Conn
    If RSPemesan.EOF Then
        Dim TambahPemesan As String
        TambahPemesan = "Insert Into Pemesan(Kodepms,Namapms,Alamat,Telepon)" & _
        "values('" & Left(CBOKodePms, 6) & "','" & NamaPms & "','" & Alamat & "','" & Telepon & "')"
        Conn.Execute (TambahPemesan)
    End If

    'simpan ke tabel MasterPO
    Dim TambahPO As String
    TambahPO = "Insert Into MasterPo(NOPO,TGLPO,JMLPO,DP,SISA,KEMBALI,BIAYA,LABA,KODEPMS,KET1,Ket2)" & _
    "values('" & NomorPO & "','" & TglPO & "','" & JumlahPO & "','" & DP & "','" & Sisa & "','" & Kembali & "','" & Biaya & "','" & Laba & "','" & Left(CBOKodePms, 6) & "','" & Ket & "', 'BELUM DIKIRIM')"
    Conn.Execute (TambahPO)

    'simpan ke tabel detailPO
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!Kode <> vbNullString Then
        Dim SimpanOrder As String
        SimpanOrder = "Insert Into DetailPO(NOPO,TANGGAL,KODEPRK,KETERANGAN,JUMLAH) " & _
        "values ('" & NomorPO & "','" & CDate(Adodc1.Recordset!Tanggal) & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!Keterangan & "','" & Adodc1.Recordset!Jumlah & "')"
        Conn.Execute (SimpanOrder)
    End If
    Adodc1.Recordset.MoveNext
    Loop
    
    'penyimpanan juga terjadi di tabel kas
    'jika DP lebih besar dari PO maka
    If DP >= JumlahPO Then
        Dim aa As String
        'simpan ke tabel kas sebagaipemasukan di kolom debet
        aa = "insert into kas (tanggal,keterangan,debet) values ('" & TglPO & "','PEMASUKAN DARI ' + '" & NamaPms & "','" & Str(DP) & "')"
        Conn.Execute aa
    'jika DP=0 maka
    ElseIf DP = 0 Or DP = "" Then
        Dim qq As String
        'simpan ke tabel kas sebagai piutang di kolom debet
        qq = "insert into kas (tanggal,keterangan,debet) values ('" & TglPO & "','PIUTANG PADA ' + '" & NamaPms & "','" & Str(Sisa) & "')"
        Conn.Execute qq
    'jika DP lebih kecil dari jumlah PO, maka
    ElseIf DP < JumlahPO Then
        Dim bb As String
        'simpan ke tabel kas sebagai uang muka di kolom debet
        bb = "insert into kas (tanggal,keterangan,debet) values ('" & TglPO & "','UANG MUKA DARI ' + '" & NamaPms & "','" & Str(DP) & "')"
        Conn.Execute bb
            
        Dim CC As String
        'simpan ke tabel kas sebagai piutang di kolom debet
        CC = "insert into kas (tanggal,keterangan,debet) values ('" & TglPO & "','PIUTANG PADA ' + '" & NamaPms & "','" & Str(Sisa) & "')"
        Conn.Execute CC
    End If
    
    'jumlah biaya dalam grid simpan ke tabel kas di kolom kredit
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
    If Adodc1.Recordset!Kode <> vbNullString Then
        Dim SimpanKasKredit As String
        SimpanKasKredit = "Insert Into Kas(TANGGAL,Keterangan,kredit) " & _
        "values ('" & CDate(Adodc1.Recordset!Tanggal) & "','" & Adodc1.Recordset!Keterangan & "','" & Adodc1.Recordset!Jumlah & "')"
        Conn.Execute (SimpanKasKredit)
    End If
    Adodc1.Recordset.MoveNext
    Loop
    
    Call BersihkanPO
    Call BersihkanPms
    Call BersihkanFooter
    Form_Activate
    Call Cetakfaktur
End If
End Sub

Sub Cetakfaktur()
    CR1.ReportFileName = App.Path & "\Faktur Order.rpt"
    CR1.WindowState = crptMaximized
    CR1.RetrieveDataFiles
    CR1.Action = 1
End Sub

Sub BersihkanPO()
JumlahPO = ""
DP = ""
Sisa = ""
Ket = ""
End Sub

Sub BersihkanPms()
NamaPms = ""
Alamat = ""
Telepon = ""
End Sub

Sub BersihkanFooter()
Biaya = ""
Laba = ""
JmlData = ""
End Sub

Private Sub CmdBatal_Click()
Form_Activate
Call Bersihkan
JmlData = ""
Call BersiTelp3ms
CBOKodePms.SetFocus
Label8 = ""
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Function TBLKosong()
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
    For i = 1 To 31
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!NO = i
        Adodc1.Recordset.Update
    Next i
    DataGrid1.Col = 1
End Function

'jumlah tidak boleh kosong
Private Sub Jumlah_KeyPress(Keyascii As Integer)
On Error Resume Next
If Keyascii = 13 Then
    If Jumlah = "" Then
        MsgBox "Jumlah masih kosong" & _
        "Isi jumlahnya dengan benar"
        Jumlah.SetFocus
        Exit Sub
    Else
        'setelah mengisi jumlah, isi kolom tgl dengan tanggal
        'isi kolom kode dengan kode dari list dan seterusnya
        Adodc1.Recordset!Tanggal = TGLTran
        Adodc1.Recordset!Kode = Left(List1, 3)
        Adodc1.Recordset!Keterangan = Keterangan
        Adodc1.Recordset!Jumlah = Jumlah
        Adodc1.Recordset.Update
        Adodc1.Recordset.MoveNext
        Call Bersihkantrans
        TGLTran.SetFocus
        DataGrid1.Col = 1
        'cari berapa jumlah baris data
        Call TotalItem
        'cari berapa total biaya
        Call TotalBiaya
        'tampilkan labanya
        Laba = Format(JumlahPO - Biaya, "###,###,###")
        'jika laba <=0 maka tampilkan pesan
        If Laba <= 0 Then
            MsgBox "Biaya lebih besar dari laba, perbaiki data dalam grid"
            DataGrid1.SetFocus
            DataGrid1.Col = 4
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Sub Bersihkantrans()
Keterangan = ""
Kode = ""
NamaPrk = ""
Jumlah = ""
End Sub

Sub Bersihkan()
JumlahPO = ""
DP = ""
Sisa = ""
Kembali = ""
Ket = ""
Laba = ""
Biaya = ""
End Sub

'fungsi untuk mencari jumlah baris data
Function TotalItem()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Item = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Kode <> vbNullString
    Item = Item + 1
    Adodc1.Recordset.MoveNext
    JmlData = Item
Loop
End Function

'fungsi untuk mencari total biaya
Function TotalBiaya()
On Error Resume Next
Adodc1.Recordset.MoveFirst
VarBiaya = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Jumlah <> vbNullString
    VarBiaya = VarBiaya + Adodc1.Recordset!Jumlah
    Adodc1.Recordset.MoveNext
    Biaya = Format(VarBiaya, "###,###,###")
Loop
End Function

'mengisi kode perkiraan dapat diambil dari list
'dengan cara dipilih lalu tekan enter
Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Kode <> Left(List1, 3) Then
        Kode = Left(List1, 3)
        Call BukaDB
        RSPerkiraan.Open "Select * from Perkiraan where KodePrk='" & Left(List1, 3) & "'", Conn
        RSPerkiraan.Requery
        If Not RSPerkiraan.EOF Then
            Kode = RSPerkiraan!KodePrk
            NamaPrk = RSPerkiraan!NamaPrk
            Jumlah.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

Sub MatiPms()
NamaPms.Enabled = False
Alamat.Enabled = False
Telepon.Enabled = False
End Sub

Sub BersiTelp3ms()
NamaPms = ""
Alamat = ""
Telepon = ""
End Sub

Sub DataPms()
    NamaPms = RSPemesan!NamaPms
    Alamat = RSPemesan!Alamat
    Telepon = RSPemesan!Telepon
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ImplodeForm(Me, 3000)
End Sub

