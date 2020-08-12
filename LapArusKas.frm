VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapArusKas 
   Caption         =   "Laporan Arus Kas"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1320
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   3000
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3000
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3000
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1250
      End
   End
End
Attribute VB_Name = "LapArusKas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'buka database
Call BukaDB
'buka tabel kas dan tampilkan tanggalnya secara terurut dari yang terkecil
RSKas.Open "Select Distinct Tanggal From Kas order By 1", Conn
RSKas.Requery
'tampilkan tanggal di combo1,2 dan 3
Do Until RSKas.EOF
    Combo1.AddItem RSKas!Tanggal
    Combo2.AddItem Format(RSKas!Tanggal, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSKas!Tanggal, "YYYY ,MM, DD")
    RSKas.MoveNext
Loop

'definisikan recordset baru
Dim RSTGL As New ADODB.Recordset
'buka tabel kas dan ambil angka bulannya saja dari field tanggal
RSTGL.Open "select distinct Tanggal from Kas", Conn
'tampilkan berulang2 bulannya dengan format 2 angka
Do While Not RSTGL.EOF
    'tampilkan angka bulan di combo4
    Combo4.AddItem Format(RSTGL!Tanggal, "MM")
    RSTGL.MoveNext
Loop

'ciptakan recordset baru
Dim RSTHN As New ADODB.Recordset
'buka tabel kas dan ambil angka tahunya saja dari field tanggal
RSTHN.Open "select distinct year(Tanggal)  as sss from Kas", Conn
'tampilkan angka tahun di combo5
Do While Not RSTHN.EOF
    Combo5.AddItem RSTHN!sss
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'Lap Harian
Private Sub Combo1_Click()
    'saring laporan dari tabel kas yang tanggalnya dipilih di combo1
    CR.SelectionFormula = "Totext({Kas.Tanggal})='" & Combo1 & "'"
    'panggil file laporan lap arus kas harian
    CR.ReportFileName = App.Path & "\Lap arus kas harian.rpt"
    'tampilkan satu layar penuh
    CR.WindowState = crptMaximized
    'jika ada perubahan isi data maka data diupdate
    CR.RetrieveDataFiles
    'tampilkan ke layar
    CR.Action = 0
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    'jika tanggal awal kosong, tampilkan pesan
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        'jika tanggal awal lebih besar dari tanggal akhir, tampilkan pesan
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Exit Sub
        'jika tgl awal = tgl akhir, tampilkan pesan
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih tanggal yang berbeda"
            Exit Sub
        End If
    End If
    'jika semua pilihan sudah benar maka, saring laporang
    'yang tgl awalnya =combo2 dan tgl akhirnya=combo3
    CR.SelectionFormula = "{Kas.Tanggal} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    'panggil file lap arus kas mingguan
    CR.ReportFileName = App.Path & "\Lap arus kas mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    'sebelum lap dipanggil cek datanya dulu
    Call BukaDB
    'buka tabel kas yg bulan dan tahunnya dipilih di combo4 dan 5
    RSKas.Open "select * from Kas where month(Tanggal)='" & Val(Combo4) & "' and year(Tanggal)='" & (Combo5) & "'", Conn
    'jika data tidak ditemukan lap tidak usah diloading, tapi munculkan pesan
    If RSKas.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    'jika datanya ada maka saring data dalam laporan
    CR.SelectionFormula = "Month({Kas.Tanggal})=" & Val(Combo4.Text) & " and Year({Kas.Tanggal})=" & Val(Combo5.Text)
    'panggil file laporannya
    CR.ReportFileName = App.Path & "\Lap arus kas bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ImplodeForm(Me, 5000)
End Sub
