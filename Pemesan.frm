VERSION 5.00
Begin VB.Form Pemesan 
   Caption         =   "Data Pemesan"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
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
   ScaleHeight     =   2115
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2250
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   900
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   900
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   900
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   4500
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   4500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   4500
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
      TabIndex        =   11
      Top             =   480
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
      TabIndex        =   10
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1245
   End
End
Attribute VB_Name = "Pemesan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
Text2.MaxLength = 30
Text3.MaxLength = 30
Text4.MaxLength = 15
Semula
End Sub

Private Sub BukaObjek()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then Ctl.Enabled = True
Next
End Sub

Private Sub TutupObjek()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then Ctl.Enabled = False
Next
End Sub

Private Sub Kosongkan()
Dim Ctl As Control
For Each Ctl In Me
    If TypeName(Ctl) = "TextBox" Or TypeName(Ctl) = "ComboBox" Then Ctl.Text = ""
Next
End Sub

Private Sub Semula()
Call Kosongkan
Call TutupObjek
CmdInput.Caption = "&Input"
CmdEdit.Caption = "&Edit"
CmdTutup.Caption = "&Tutup"
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
On Error Resume Next
Text2 = RSPemesan!NamaPms
Text3 = RSPemesan!Alamat
Text4 = RSPemesan!Telepon
End Sub

Private Sub AutoNomor()
Call BukaDB
RSPemesan.Open ("select * from Pemesan Where KodePms In(Select Max(KodePms)From Pemesan)Order By KodePms Desc"), Conn
Dim Urutan As String * 6
Dim Hitung As Long
With RSPemesan
    If .EOF Then
        Urutan = "PMS" + "001"
        Combo1 = Urutan
    Else
        Hitung = Right(!KodePms, 3) + 1
        Urutan = "PMS" + Right("000" & Hitung, 3)
    End If
    Combo1 = Urutan
End With
End Sub

Private Sub CmdInput_Click()
If CmdInput.Caption = "&Input" Then
    CmdInput.Caption = "&Simpan"
    CmdEdit.Enabled = False
    CmdHapus.Enabled = False
    CmdTutup.Caption = "&Batal"
    BukaObjek
    Kosongkan
    Combo1.SetFocus
    Call AutoNomor
    Combo1.Enabled = False
    Text2.SetFocus
Else
    If Combo1 = "" Or Text2 = "" Or Text3 = "" Then
        MsgBox "Kode, Nama dan alamat wajib diisi...!"
    Else
        Dim SQLTambah As String
        SQLTambah = "Insert Into Pemesan (KodePms,NamaPms,alamat,Telepon) values " & _
        "('" & Combo1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
        Conn.Execute SQLTambah
        Semula
    End If
End If
End Sub

Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    CmdInput.Enabled = False
    CmdEdit.Caption = "&Simpan"
    CmdHapus.Enabled = False
    CmdTutup.Caption = "&Batal"
    BukaObjek
    Call BukaDB
    RSPemesan.Open "select * from pemesan", Conn
    Combo1.Clear
    Do Until RSPemesan.EOF
        Combo1.AddItem RSPemesan!KodePms
        RSPemesan.MoveNext
    Loop
    Combo1.SetFocus
Else
    If Text2 = "" Or Text3 = "" Then
        MsgBox "Nama dan alamat wajib diisi"
    Else
        Dim SQLEdit As String
        SQLEdit = "Update Pemesan Set " & _
        "NamaPms= '" & Text2 & "', " & _
        "Alamat='" & Text3 & "', " & _
        "Telepon='" & Text4 & "' " & _
        "where KodePms='" & Combo1 & "'"
        Conn.Execute SQLEdit
        Semula
    End If
End If
End Sub

Private Sub CmdHapus_Click()
If CmdHapus.Caption = "&Hapus" Then
    CmdInput.Enabled = False
    CmdEdit.Enabled = False
    CmdTutup.Caption = "&Batal"
    Kosongkan
    BukaObjek
    Call BukaDB
    RSPemesan.Open "select * from pemesan", Conn
    Combo1.Clear
    Do Until RSPemesan.EOF
        Combo1.AddItem RSPemesan!KodePms
        RSPemesan.MoveNext
    Loop
    Combo1.SetFocus
End If
End Sub

Private Sub CmdTutup_Click()
If CmdTutup.Caption = "&Tutup" Then Unload Me
If CmdTutup.Caption = "&Batal" Then Call Semula
End Sub

Function CariData()
Call BukaDB
RSPemesan.Open "Select * From Pemesan where KodePms='" & Combo1 & "'", Conn
End Function

Private Sub Combo1_Click()
Call BukaDB
Call CariData
Call TampilkanData
End Sub

Private Sub Combo1_Keypress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Combo1) <> 6 Then
        MsgBox "Kode Harus 6 Digit, Contoh PMS001"
        Combo1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
        If Not RSPemesan.EOF Then
            TampilkanData
            MsgBox "Kode Pemesan Sudah Ada"
            Kosongkan
            Combo1.SetFocus
        Else
            Text2.SetFocus
        End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
        If Not RSPemesan.EOF Then
            TampilkanData
            Combo1.Enabled = False
            Text2.SetFocus
        Else
            MsgBox "Kode Pemesan Tidak Ada"
            Combo1 = ""
            Combo1.SetFocus
        End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
        If Not RSPemesan.EOF Then
            Call TampilkanData
            Pesan = MsgBox("Yakin akan dihapus..?", vbYesNo)
            If Pesan = vbYes Then
                Dim hapus As String
                hapus = "Delete from Pemesan where kodepms='" & Combo1 & "'"
                Conn.Execute hapus
                Call Semula
            Else
                Call Semula
            End If
        Else
            MsgBox "Kode pemesan tidak ditemukan"
            Combo1.SetFocus
            Exit Sub
        End If
    End If
End If
End Sub

'nama
Private Sub Text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

'alamat
Private Sub Text3_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text4.SetFocus
End Sub

'telepon
Private Sub Text4_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
Me.Top = Me.Top + 40
Me.Move Me.Left, Me.Top
DoEvents
Loop Until Me.Top > Screen.Height - 500
End Sub

