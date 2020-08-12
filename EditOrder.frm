VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EditOrder 
   Caption         =   "Edit Kesalahan Order"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan Setelah Diedit"
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2000
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   350
      Left            =   2160
      TabIndex        =   3
      Top             =   5760
      Width           =   750
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3000
      TabIndex        =   2
      Top             =   5760
      Width           =   750
   End
   Begin VB.ComboBox CboNOPO 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "EditOrder.frx":0000
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5027
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   120
      Top             =   6240
      Visible         =   0   'False
      Width           =   3105
      _ExtentX        =   5477
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Telp3 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   36
      Top             =   2280
      Width           =   3800
   End
   Begin VB.Label Telp2 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   35
      Top             =   1920
      Width           =   3800
   End
   Begin VB.Label Telp1 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   34
      Top             =   1560
      Width           =   3800
   End
   Begin VB.Label Alamat1 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   33
      Top             =   840
      Width           =   3800
   End
   Begin VB.Label Alamat2 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   32
      Top             =   1200
      Width           =   3800
   End
   Begin VB.Label KodePms 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   31
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label NamaPms 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4680
      TabIndex        =   30
      Top             =   480
      Width           =   3800
   End
   Begin VB.Label JMLPO 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   29
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label TglPO 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   28
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label DP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   27
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor PO"
      Height          =   345
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal PO"
      Height          =   345
      Left            =   120
      TabIndex        =   25
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah "
      Height          =   345
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label JmlData 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   3840
      TabIndex        =   23
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label Biaya 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   7080
      TabIndex        =   22
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Pemesan"
      Height          =   345
      Left            =   3360
      TabIndex        =   21
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Pemesan"
      Height          =   345
      Left            =   3360
      TabIndex        =   20
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat 2"
      Height          =   345
      Left            =   3360
      TabIndex        =   19
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat 1"
      Height          =   345
      Left            =   3360
      TabIndex        =   18
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telp Rumah"
      Height          =   345
      Left            =   3360
      TabIndex        =   17
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telp. Kantor"
      Height          =   345
      Left            =   3360
      TabIndex        =   16
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No. HP"
      Height          =   345
      Left            =   3360
      TabIndex        =   15
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Laba 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   14
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Laba"
      Height          =   345
      Left            =   5760
      TabIndex        =   13
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Muka"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Keterangan"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Ket 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Label Sisa 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Biaya"
      Height          =   345
      Left            =   7080
      TabIndex        =   7
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Label Kembali 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kembali"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1245
   End
End
Attribute VB_Name = "EditOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Activate()
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBAKN.mdb"
DT.RecordSource = "EditTrans"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
Call TBLKosong
'CmdSimpan.Enabled = False
CboNOPO.SetFocus

Call BukaDB
RSMasterPO.Open "Select * from Masterpo ", Conn
CboNOPO.Clear
Do Until RSMasterPO.EOF
    CboNOPO.AddItem RSMasterPO!NOPO
    RSMasterPO.MoveNext
Loop
Conn.Close
End Sub

Private Sub Form_Load()
    DataGrid1.Col = 1
    'CmdSimpan.Enabled = False
End Sub

Function TBLKosong()
If DT.Recordset.RecordCount > 0 Then
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        DT.Recordset.Delete
        DT.Recordset.MoveNext
    Loop
End If
End Function

Private Sub CBONOPO_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
If KeyAscii = 13 Then
    Call BukaDB
    RSMasterPO.Open "Select * from Masterpo where NOPO='" & CboNOPO & "'", Conn
    RSMasterPO.Requery
    If RSMasterPO.EOF Then
        MsgBox "NO Masterpo tidak terdaftar"
        CboNOPO.SetFocus
        Exit Sub
    Else
        DataGrid1.SetFocus
    End If
End If

End Sub

Private Sub CBONOPO_Click()
Call BukaDB
RSMasterPO.Open "Select * from Masterpo where NOPO='" & CboNOPO & "'", Conn
RSMasterPO.Requery
If Not RSMasterPO.EOF Then
    'TglMintakrm = CDate(RSMasterPO!TglMintakrm)
    'Total = Format(RSMasterPO!TotalHrg, "###,###,###")
    'Sisa = Format(RSMasterPO!Sisa, "###,###,###")
    'DP = Format(RSMasterPO!DP, "###,###,###")
    
    Biaya = Format(RSMasterPO!Biaya, "###,###,###")
    Laba = Format(RSMasterPO!Laba, "###,###,###")
    TGLPO = RSMasterPO!TGLPO
    JMLPO = Format(RSMasterPO!JMLPO, "###,###,###")
    DP = Format(RSMasterPO!DP, "###,###,###")
    Sisa = Format(RSMasterPO!Sisa, "###,###,###")
    Kembali = RSMasterPO!Kembali
    Ket = RSMasterPO!ket1
    KodePms = RSMasterPO!KodePms


    Dim RS As New ADODB.Recordset
    RS.Open "select * from detailpo where nopo='" & CboNOPO & "'", Conn
    Call TBLKosong
    RS.MoveFirst
    NO = 0
    Do While Not RS.EOF
        NO = NO + 1
        DT.Recordset.AddNew
        DT.Recordset!NO = NO
        DT.Recordset!Tanggal = RS!Tanggal
        DT.Recordset!Kode = RS!KodePrk
        DT.Recordset!Keterangan = RS!Keterangan
        DT.Recordset!Jumlah = RS!Jumlah
        DT.Recordset.Update
        RS.MoveNext
    Loop
Else
    MsgBox "NO Order tidak terdaftar"
    CboNOPO.SetFocus
    Exit Sub
End If
End Sub

Private Sub Kodepms_Change()
Call BukaDB
RSPemesan.Open "Select * from Pemesan where kodepms='" & KodePms & "'", Conn
RSPemesan.Requery
If Not RSPemesan.EOF Then
    NamaPms = RSPemesan!NamaPms
    Alamat1 = RSPemesan!Alamat1
    Alamat2 = RSPemesan!Alamat2
    Telp1 = RSPemesan!Telepon1
    Telp2 = RSPemesan!Telepon2
    Telp3 = RSPemesan!telepon3
End If
End Sub

Private Sub CmdSimpan_Click()
    If CboNOPO = "" Then
        MsgBox "NO order tidak boleh kosong"
        Exit Sub
    End If
    
    DT.Recordset.MoveFirst
    Do While Not DT.Recordset.EOF
        Dim SimpanDetailPO As String
        SimpanDetailPO = "Update detailpo set Tanggal='" & DT.Recordset!Tanggal & "', kodeprk='" & DT.Recordset!Kode & "',keterangan='" & DT.Recordset!Keterangan & "', jumlah='" & DT.Recordset!Jumlah & "' where nopo='" & CboNOPO & "' and keterangan='" & DT.Recordset!Keterangan & "'"
        Conn.Execute (SimpanDetailPO)
    DT.Recordset.MoveNext
    Loop
       
    bersihkan
    Form_Activate
End Sub

Private Sub CmdBatal_Click()
Call bersihkan
End Sub

Sub bersihkan()
CboNOPO = ""
TGLPO = ""
JMLPO = ""
DP = ""
Sisa = ""
Kembali = ""
Ket = ""
Biaya = ""
Laba = ""

KodePms = ""
NamaPms = ""
Alamat1 = ""
Alamat2 = ""
Telp1 = ""
Telp2 = ""
Telp3 = ""

Call TBLKosong
End Sub
