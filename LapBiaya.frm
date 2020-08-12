VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LapBiaya 
   Caption         =   "Laporan Biaya - Biaya"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
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
   ScaleHeight     =   5175
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox List1 
      Height          =   3885
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tahun"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bulan"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1250
   End
End
Attribute VB_Name = "LapBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'pola program ini sama dengan laporan arus kas

Private Sub Form_Load()
Call BukaDB
Dim RS1 As New ADODB.Recordset
RS1.Open "select distinct perkiraan.kodeprk,perkiraan.namaprk from perkiraan,detailpo where perkiraan.kodeprk=detailpo.kodeprk", Conn
List1.Clear
Do While Not RS1.EOF
    List1.AddItem RS1!KodePrk & Space(3) & RS1!NamaPrk
    RS1.MoveNext
Loop

Dim RSTGL As New ADODB.Recordset
RSTGL.Open "select distinct tanggal from detailpo", Conn
Do While Not RSTGL.EOF
    Combo1.AddItem Format(RSTGL!Tanggal, "MM")
    RSTGL.MoveNext
Loop

Dim RSTHN As New ADODB.Recordset
RSTHN.Open "select distinct year(tanggal)  as sss from detailpo", Conn
Do While Not RSTHN.EOF
    Combo2.AddItem RSTHN!sss
    RSTHN.MoveNext
Loop
Conn.Close
End Sub

Private Sub List1_Click()
If Combo1 = "" Or Combo2 = "" Then
    MsgBox "Bulan dan Tahun tidak boleh kosong"
    Combo1.SetFocus
    Exit Sub
End If

Call BukaDB
Dim RS2 As New ADODB.Recordset
RS2.Open "select * from detailpo where month(tanggal)='" & Val(Combo1) & "' and year(tanggal)='" & (Combo2) & "' and kodeprk='" & Left(List1, 3) & "'", Conn
If RS2.EOF Then
    MsgBox "Data tidak ditemukan"
    Combo1.SetFocus
    Exit Sub
Else
    If Left(List1, 3) = "401" Then
        CR.SelectionFormula = "{Detailpo.Kodeprk}='" & Left(List1, 3) & "' and Month({detailPO.tanggal})=" & Val(Combo1.Text) & " and Year({Detailpo.Tanggal})=" & Val(Combo2.Text)
        CR.ReportFileName = App.Path & "\Lap oprs kendaraan.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Else
        CR.SelectionFormula = "{Detailpo.Kodeprk}='" & Left(List1, 3) & "' and Month({detailPO.tanggal})=" & Val(Combo1.Text) & " and Year({Detailpo.Tanggal})=" & Val(Combo2.Text)
        CR.ReportFileName = App.Path & "\Lap umum.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    End If
End If
End Sub

