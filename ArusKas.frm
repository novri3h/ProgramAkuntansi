VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ArusKas 
   Caption         =   "Laporan Arus Kas"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ArusKas.frx":0000
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "NO"
         Caption         =   "NO"
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
         DataField       =   "TANGGAL"
         Caption         =   "TANGGAL"
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
         DataField       =   "KETERANGAN"
         Caption         =   "KETERANGAN"
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
         DataField       =   "DEBET"
         Caption         =   "DEBET"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "KREDIT"
         Caption         =   "KREDIT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "SALDO"
         Caption         =   "SALDO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
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
            Alignment       =   2
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal / Bulan"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label TTLDebet 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   4200
      TabIndex        =   5
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label TTLKredit 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   4
      Top             =   5160
      Width           =   1500
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   5760
      TabIndex        =   3
      Top             =   5520
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALDO"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   5520
      Width           =   1500
   End
End
Attribute VB_Name = "ArusKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBAKN.mdb"
Adodc1.RecordSource = "TRArusKas"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Call TBLKosong
End Sub

Private Sub Form_Load()
Call BukaDB
RSKas.Open "Select distinct tanggal from kas", Conn
Combo1.Clear
Do While Not RSKas.EOF
    Combo1.AddItem RSKas!Tanggal
    RSKas.MoveNext
Loop
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

Private Sub Combo1_Click()
On Error Resume Next
Call BukaDB
RSKas.Open "Select * from kas where cdate(tanggal)='" & Combo1 & "'", Conn
If Not RSKas.EOF Then
    Call TBLKosong
    RSKas.MoveFirst
    NO = 0
    Saldo = 0
    Do While Not RSKas.EOF
        NO = NO + 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!NO = NO
        Adodc1.Recordset!Tanggal = RSKas!Tanggal
        Adodc1.Recordset!Keterangan = RSKas!Keterangan
        Adodc1.Recordset!debet = RSKas!debet
        Adodc1.Recordset!kredit = RSKas!kredit
        If Adodc1.Recordset.RecordCount = 1 Then
            Adodc1.Recordset!Saldo = RSKas!debet - RSKas!kredit
        Else
            Adodc1.Recordset.MovePrevious
            Saldo = Adodc1.Recordset!Saldo
            Adodc1.Recordset.MoveNext
            Adodc1.Recordset!Saldo = (Saldo + Adodc1.Recordset!debet) - Adodc1.Recordset!kredit
        End If
        RSKas.MoveNext
        Call TotalDebet
        Call TotalKredit
    Loop
    Saldo = Str(TTLDebet) - Str(TTLKredit)
    Saldo = Format(Saldo, "###,###,###")
Else
    MsgBox "Tanggal tidak valid"
    Combo1.SetFocus
    Exit Sub
End If
End Sub

Function TotalDebet()
On Error Resume Next
Adodc1.Recordset.MoveFirst
debet = 0
Do While Not Adodc1.Recordset.EOF
    debet = debet + Adodc1.Recordset!debet
    Adodc1.Recordset.MoveNext
    TTLDebet = Format(debet, "###,###,###")
Loop
End Function

Function TotalKredit()
On Error Resume Next
Adodc1.Recordset.MoveFirst
kredit = 0
Do While Not Adodc1.Recordset.EOF
    kredit = kredit + Adodc1.Recordset!kredit
    Adodc1.Recordset.MoveNext
    TTLKredit = Format(kredit, "###,###,###")
Loop
End Function


