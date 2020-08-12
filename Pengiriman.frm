VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pengiriman 
   Caption         =   "Pengiriman Barang"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8040
      Top             =   5400
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
      TabIndex        =   4
      Top             =   5400
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
      TabIndex        =   3
      Top             =   5400
      Width           =   750
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
      TabIndex        =   2
      Top             =   5400
      Width           =   750
   End
   Begin VB.ComboBox CboNOPO 
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
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pengiriman.frx":0000
      Height          =   2850
      Left            =   120
      TabIndex        =   0
      Top             =   2400
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3360
      Top             =   1920
      Visible         =   0   'False
      Width           =   5115
      _ExtentX        =   9022
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
   Begin VB.Label Alamat 
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
      Left            =   4680
      TabIndex        =   27
      Top             =   1200
      Width           =   3795
   End
   Begin VB.Label Telepon 
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
      Left            =   4680
      TabIndex        =   26
      Top             =   1560
      Width           =   3795
   End
   Begin VB.Label KodePms 
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
      Left            =   4680
      TabIndex        =   25
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label NamaPms 
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
      Left            =   4680
      TabIndex        =   24
      Top             =   840
      Width           =   3795
   End
   Begin VB.Label JMLPO 
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
      Left            =   1440
      TabIndex        =   23
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label TglPO 
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
      Left            =   4680
      TabIndex        =   22
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label DP 
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
      Left            =   1440
      TabIndex        =   21
      Top             =   840
      Width           =   1800
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
      Height          =   345
      Left            =   120
      TabIndex        =   20
      Top             =   120
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
      Left            =   3360
      TabIndex        =   19
      Top             =   120
      Width           =   1245
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
      Left            =   120
      TabIndex        =   18
      Top             =   480
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
      Left            =   6480
      TabIndex        =   17
      Top             =   5400
      Width           =   1500
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
      Left            =   3360
      TabIndex        =   16
      Top             =   840
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
      Left            =   3360
      TabIndex        =   15
      Top             =   480
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
      Left            =   3360
      TabIndex        =   14
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat "
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Muka"
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
      TabIndex        =   12
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sisa"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1245
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
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1245
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
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   1800
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1200
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
      Left            =   4920
      TabIndex        =   7
      Top             =   5400
      Width           =   1500
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
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
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
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1245
   End
End
Attribute VB_Name = "Pengiriman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Me.Width = Me.Width + 100
Tengah
If Me.Width >= 8715 Then
    Timer1.Enabled = False
    Tengah
End If
End Sub

Public Sub Tengah()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub DataGrid1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBAKN.mdb"
Adodc1.RecordSource = "EditTrans"
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call TBLKosong
CboNOPO.SetFocus
Call BukaDB
RSMasterPO.Open "Select * from Masterpo where ket2='BELUM DIKIRIM'", Conn
CboNOPO.Clear
Do Until RSMasterPO.EOF
    CboNOPO.AddItem RSMasterPO!NOPO
    RSMasterPO.MoveNext
Loop
Conn.Close
End Sub

Private Sub Form_Load()
Me.Width = 10
End Sub

Function TBLKosong()
If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
End If
End Function

Private Sub CBONOPO_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
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
   
    Biaya = Format(RSMasterPO!Biaya, "###,###,###")
    'Laba = Format(RSMasterPO!Laba, "###,###,###")
    TglPO = RSMasterPO!TglPO
    JMLPO = Format(RSMasterPO!JMLPO, "###,###,###")
    
    If RSMasterPO!DP = 0 Then
        DP = 0
    Else
        DP = Format(RSMasterPO!DP, "###,###,###")
    End If
    
    If RSMasterPO!Sisa = 0 Then
        Sisa = 0
    Else
        Sisa = Format(RSMasterPO!Sisa, "###,###,###")
    End If
        
    If RSMasterPO!Kembali = 0 Then
        Kembali = 0
    Else
        Kembali = Format(RSMasterPO!Kembali, "###,###,###")
    End If
    
    Ket = RSMasterPO!ket1
    KodePms = RSMasterPO!KodePms

    Dim RS As New ADODB.Recordset
    RS.Open "select * from detailpo where nopo='" & CboNOPO & "'", Conn
    Call TBLKosong
    RS.MoveFirst
    NO = 0
    Do While Not RS.EOF
        NO = NO + 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!NO = NO
        Adodc1.Recordset!Tanggal = RS!Tanggal
        Adodc1.Recordset!Kode = RS!KodePrk
        Adodc1.Recordset!Keterangan = RS!Keterangan
        Adodc1.Recordset!Jumlah = RS!Jumlah
        Adodc1.Recordset.Update
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
    Alamat = RSPemesan!Alamat
    Telepon = RSPemesan!Telepon
End If
End Sub

Private Sub CmdSimpan_Click()
    If CboNOPO = "" Then
        MsgBox "NO order tidak boleh kosong"
        Exit Sub
    End If
    Dim aa As String
    aa = "update masterpo set ket2='SUDAH DIKIRIM' where nopo='" & CboNOPO & "'"
    Conn.Execute aa
    Call Bersihkan
    Form_Activate
End Sub

Private Sub CmdBatal_Click()
Call Bersihkan
End Sub

Sub Bersihkan()
CboNOPO = ""
TglPO = ""
JMLPO = ""
DP = ""
Sisa = ""
Kembali = ""
Ket = ""
Biaya = ""
Laba = ""
KodePms = ""
NamaPms = ""
Alamat = ""
Telepon = ""
Call TBLKosong
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ImplodeForm(Me, 5000)
End Sub
