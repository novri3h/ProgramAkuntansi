VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapOrder 
   Caption         =   "Laporan Pesanan"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   2880
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo8 
      Height          =   345
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Width           =   1500
   End
   Begin VB.ComboBox Combo7 
      Height          =   345
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Frame Frame4 
      Caption         =   "Nomor Pesanan (Order)"
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   3000
      Begin VB.ComboBox Combo6 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nomor"
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   3000
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Akhir"
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal Awal"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   3000
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   3240
      TabIndex        =   13
      Top             =   1200
      Width           =   3000
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1250
      End
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dikirim / Belum"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   19
      Top             =   2640
      Width           =   1245
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lunas / Belum"
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1245
   End
End
Attribute VB_Name = "LapOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB
'buka tabel masterpo
RSMasterPO.Open "Select * From MasterPO order By 1", Conn
RSMasterPO.Requery
Do Until RSMasterPO.EOF
    'tampilkan nomor po di combo6
    Combo6.AddItem RSMasterPO!NOPO
    RSMasterPO.MoveNext
Loop
Conn.Close

Call BukaDB
'buka tabel masterpo dan tampilkan field Ket1
RSMasterPO.Open "Select distinct Ket1 From MasterPO", Conn
RSMasterPO.Requery
Do Until RSMasterPO.EOF
    'tampilkan ket1 di combo7
    Combo7.AddItem RSMasterPO!ket1
    RSMasterPO.MoveNext
Loop
Conn.Close

Call BukaDB
RSMasterPO.Open "Select distinct Ket2 From MasterPO", Conn
RSMasterPO.Requery
Do Until RSMasterPO.EOF
    'tampilkan field ket2 di combo8
    Combo8.AddItem RSMasterPO!ket2
    RSMasterPO.MoveNext
Loop
Conn.Close


Call BukaDB
'buka tabel masterpo dan ambil tanggalnya saja
RSMasterPO.Open "Select Distinct TGLPO From MasterPO order By 1", Conn
RSMasterPO.Requery
Do Until RSMasterPO.EOF
    'tampilkan tglpo di combo1,2 dan 3
    Combo1.AddItem RSMasterPO!TglPO
    Combo2.AddItem Format(RSMasterPO!TglPO, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSMasterPO!TglPO, "YYYY ,MM, DD")
    RSMasterPO.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct tGLpo from MASTERPO", Conn
Do While Not RSTGL.EOF
    'tampilkan tglpo berupa angka bulan di combo4
    Combo4.AddItem Format(RSTGL!TglPO, "MM")
    RSTGL.MoveNext
Loop
Conn.Close

Call BukaDB
Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(tGLpo)  as sss from MASTERPO", Conn
Do While Not RSTHN.EOF
    'tampilkan angka tahun di combo5
    Combo5.AddItem RSTHN!sss
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'lap per nomor po
Private Sub Combo6_Click()
    CR.SelectionFormula = "{MasterPO.NOPO}='" & Combo6 & "'"
    CR.ReportFileName = App.Path & "\Faktur order.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({MasterPO.TGLPO})='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Lap Order Harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "TGLPO awal kosong", , "Informasi"
        Combo2.SetFocus
        Exit Sub
    Else
        If Combo3 < Combo2 Or Combo2 > Combo3 Then
            MsgBox "Tanggal terbalik"
            Exit Sub
        ElseIf Combo3 = Combo2 Then
            MsgBox "pilih tanggal yang berbeda"
            Exit Sub
        End If
    End If
    CR.SelectionFormula = "{MasterPO.TGLPO} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Lap Order Mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call BukaDB
    RSMasterPO.Open "select * from MasterPO where month(TGLPO)='" & Val(Combo4) & "' and year(TGLPO)='" & (Combo5) & "'", Conn
    If RSMasterPO.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    
    CR.SelectionFormula = "Month({MasterPO.TGLPO})=" & Val(Combo4.Text) & " and Year({MasterPO.TGLPO})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap Order Bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo7_Click()
    CR.SelectionFormula = "{MASTERPO.KET1}='" & Combo7 & "'"
    CR.ReportFileName = App.Path & "\Lap KET1.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo8_Click()
    CR.SelectionFormula = "{MASTERPO.KET2}='" & Combo8 & "'"
    CR.ReportFileName = App.Path & "\Lap KET2.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

