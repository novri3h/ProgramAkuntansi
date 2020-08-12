VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Kas 
   Caption         =   "SIMPAN KAS"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3180
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
   ScaleHeight     =   1245
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Jumlah 
      Alignment       =   1  'Right Justify
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1500
   End
   Begin MSComCtl2.DTPicker Tanggal 
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
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
      Format          =   94240769
      CurrentDate     =   39565
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "Kas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'buka database
Call BukaDB
'tanggal diambil dari sistem komputer
Tanggal = Date
End Sub

Private Sub Jumlah_KeyPress(Keyascii As Integer)
'jika menekan enter maka
If Keyascii = 13 Then
    'jika jumlah kosong atau jumlah nol maka
    If Jumlah = "" Or Jumlah = 0 Then
        'tampilkan pesan
        MsgBox "jumlah pemasukan masih kosong"
        Jumlah.SetFocus
        Exit Sub
    Else
        'jika jumlah telah diisi maka ubah formatnya
        Jumlah = Format(Jumlah, "###,###,###")
        'tampilkan pesan
        Pesan = MsgBox("Data sudah benar..?", vbYesNo)
        'jika pesan dibawaj YES maka
        If Pesan = vbYes Then
            Dim Simpankas As String
            'simpan data ke tabel kas, keterangan diambil dari caption form
            Simpankas = "insert into kas (tanggal,Keterangan,DEBET) values ('" & Tanggal & "','" & Kas.Caption & "','" & Jumlah & "')"
            Conn.Execute (Simpankas)
            Jumlah = ""
            Tanggal.SetFocus
        End If
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
Me.Top = Me.Top + 40
Me.Move Me.Left, Me.Top
DoEvents
Loop Until Me.Top > Screen.Height - 500
End Sub

