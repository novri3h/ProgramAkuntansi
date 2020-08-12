VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form ArusKas1 
   Caption         =   "Laporan Arus Kas"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   1320
      Top             =   960
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
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
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
         Height          =   315
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
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
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
Attribute VB_Name = "ArusKas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call BukaDB

RSKas.Open "Select Distinct Tanggal From Kas order By 1", Conn
RSKas.Requery
Do Until RSKas.EOF
    Combo1.AddItem RSKas!Tanggal
    Combo2.AddItem Format(RSKas!Tanggal, "YYYY ,MM, DD")
    Combo3.AddItem Format(RSKas!Tanggal, "YYYY ,MM, DD")
    RSKas.MoveNext
Loop

Call BukaDB
Dim rstgl As New ADODB.Recordset
rstgl.Open "select distinct Tanggal from Kas", Conn
Do While Not rstgl.EOF
    Combo4.AddItem Format(rstgl!Tanggal, "MM")
    rstgl.MoveNext
Loop
Conn.Close

Call BukaDB
Dim rsThn As New ADODB.Recordset
rsThn.Open "select distinct year(Tanggal)  as sss from Kas", Conn
Do While Not rsThn.EOF
    Combo5.AddItem rsThn!sss
    rsThn.MoveNext
Loop
Conn.Close

End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
If Combo1 = "" Or Keyascii = 27 Then Unload Me
End Sub

'Lap Harian
Private Sub Combo1_Click()
    CR.SelectionFormula = "Totext({Kas.Tanggal})='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\Lap arus kas harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
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
    CR.SelectionFormula = "{Kas.Tanggal} in date (" & Combo2.Text & ") to date (" & Combo3.Text & ")"
    CR.ReportFileName = App.Path & "\Lap arus kas mingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Combo4_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

'Lap Bulanan
Private Sub Combo5_Click()
    Call BukaDB
    RSKas.Open "select * from Kas where month(Tanggal)='" & Val(Combo4) & "' and year(Tanggal)='" & (Combo5) & "'", Conn
    If RSKas.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    
    CR.SelectionFormula = "Month({Kas.Tanggal})=" & Val(Combo4.Text) & " and Year({Kas.Tanggal})=" & Val(Combo5.Text)
    CR.ReportFileName = App.Path & "\Lap arus kas bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

