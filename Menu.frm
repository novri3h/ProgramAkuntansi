VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Menu 
   Caption         =   " Menu Utama Program Akuntansi PT XXX  ***   "
   ClientHeight    =   7830
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   16200
   ControlBox      =   0   'False
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   1440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":201D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20622
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":21318
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2176A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":21BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2200E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":22460
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":228B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":22D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":23156
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   840
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnfile 
      Caption         =   "File Master"
      Begin VB.Menu mnlogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnubahpwd 
         Caption         =   "Ubah Password"
      End
      Begin VB.Menu Mnkeluar1 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu mnmaster 
      Caption         =   "Master"
      Begin VB.Menu mnuser 
         Caption         =   "User"
      End
      Begin VB.Menu mnsimpankas 
         Caption         =   "Simpan Kas"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnperkiraan 
         Caption         =   "Perkiraan Biaya"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnPemesan 
         Caption         =   "Pemesan"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnorder 
         Caption         =   "Order"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnpengiriman 
         Caption         =   "Pengiriman"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlaporder 
         Caption         =   "Order"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnlaparuskas 
         Caption         =   "Arus Kas"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnlapbiaya 
         Caption         =   "Biaya - Biaya"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnbukubesar 
         Caption         =   "Buku Besar"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu Mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Bergerak As Integer
Dim Teks As String

Private Sub Form_Load()
Teks = Me.Caption
End Sub

Private Sub Mnkeluar1_Click()
Pesan = MsgBox("Yakin akan mengakhiri program ini", vbYesNo)
If Pesan = vbYes Then End

End Sub

Private Sub mnlogin_Click()
Login.Show
End Sub

Private Sub mnPemesan_Click()
Pemesan.Show vbModal
End Sub

Private Sub mnubahpwd_Click()
GantiPass.Show
End Sub

Private Sub mnuser_Click()
Kasir.Show
End Sub

Private Sub Timer1_Timer()
Me.Caption = Bergerak
Teks = Right(Teks, Len(Teks) - 1) & Left(Teks, 1)
Me.Caption = Teks
End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Pesan = MsgBox("Yakin akan keluar dari program ini..?", vbYesNo)
    If Pesan = vbYes Then
        Call ImplodeForm(Me, 1000)
        End
    End If
End If
End Sub

Private Sub mnbukubesar_Click()
    CR.ReportFileName = App.Path & "\Lap Buku Besar.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnkeluar_Click()
Pesan = MsgBox("Yakin akan mengakhiri program ini", vbYesNo)
If Pesan = vbYes Then End
End Sub

Private Sub mnlaparuskas_Click()
LapArusKas.Show vbModal
End Sub

Private Sub mnlapbiaya_Click()
LapBiaya.Show vbModal
End Sub

Private Sub mnlaporder_Click()
LapOrder.Show vbModal
End Sub

Private Sub mnorder_Click()
Order.Show vbModal
End Sub

Private Sub mnsimpankas_Click()
Kas.Show vbModal
End Sub

Private Sub mnpengiriman_Click()
Pengiriman.Show vbModal
End Sub

Private Sub mnperkiraan_Click()
Perkiraan.Show vbModal
End Sub

Sub ceksaldo()
Call BukaDB
RSKas.Open "select * from kas where saldo>0", Conn
RSKas.Requery
If RSKas.EOF Then
    Pesan = MsgBox("kas masih kosong, simpan uang kas dulu", vbYesNo)
    If Pesan = vbYes Then
        Kas.Show
    Else
        End
    End If
End If

End Sub

Private Sub mnsql_Click()
UjiSQL.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "F1"
        Kas.Show vbModal
    Case "F2"
        Perkiraan.Show vbModal
    Case "F3"
        Pemesan.Show vbModal
    Case "F4"
        Order.Show vbModal
    Case "F5"
        Pengiriman.Show vbModal
    Case "F6"
        LapOrder.Show vbModal
    Case "F7"
        LapArusKas.Show vbModal
    Case "F8"
        LapBiaya.Show vbModal
    Case "F9"
        CR.ReportFileName = App.Path & "\Lap Buku Besar.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "ESC"
        Pesan = MsgBox("Yakin akan keluar dari program ini..?", vbYesNo)
        If Pesan = vbYes Then
            Call ImplodeForm(Me, 5000)
            End
        End If
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ImplodeForm(Me, 1000)
End Sub
