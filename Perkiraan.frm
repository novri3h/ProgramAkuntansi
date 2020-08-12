VERSION 5.00
Begin VB.Form Perkiraan 
   Caption         =   "Data Perkiraan"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
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
   ScaleHeight     =   1395
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1560
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   4185
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Perkiraan"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1400
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Perkiraan"
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1400
   End
End
Attribute VB_Name = "Perkiraan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
Text1.MaxLength = 30
Semula
End Sub

Private Sub KosongkanText()
    Combo1 = ""
    Text1 = ""
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text1.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text1.Enabled = False
End Sub

Private Sub Semula()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Combo1.SetFocus
    Else
        If Combo1 = "" Or Text1 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Perkiraan (KodePrk,NamaPrk) values ('" & Combo1 & "','" & Text1 & "')"
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
        SiapIsi
        Call BukaDB
        RSPerkiraan.Open "select * from Perkiraan", Conn
        Combo1.Clear
        Do Until RSPerkiraan.EOF
            Combo1.AddItem RSPerkiraan!KodePrk
            RSPerkiraan.MoveNext
        Loop
        Combo1.SetFocus
    Else
        If Text1 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Perkiraan Set NamaPrk= '" & Text1 & "' where KodePrk='" & Combo1 & "'"
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
        SiapIsi
        Call BukaDB
        RSPerkiraan.Open "select * from Perkiraan", Conn
        Combo1.Clear
        Do Until RSPerkiraan.EOF
            Combo1.AddItem RSPerkiraan!KodePrk
            RSPerkiraan.MoveNext
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
    RSPerkiraan.Open "Select * From Perkiraan where KodePrk='" & Combo1 & "'", Conn
End Function

Private Sub Combo1_Click()
Call BukaDB
Call CariData
Text1 = RSPerkiraan!NamaPrk
End Sub


Private Sub Combo1_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
    If Len(Combo1) <> 3 Then
        MsgBox "Kode Harus 3 Digit"
        Combo1.SetFocus
        Exit Sub
    Else
        Text1.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
        If Not RSPerkiraan.EOF Then
            Text1 = RSPerkiraan!NamaPrk
            MsgBox "Kode Perkiraan Sudah Ada"
            KosongkanText
            Combo1.SetFocus
        Else
            Text1.SetFocus
        End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
        If Not RSPerkiraan.EOF Then
            Text1 = RSPerkiraan!NamaPrk
            Combo1.Enabled = False
            Text1.SetFocus
        Else
            MsgBox "Kode Perkiraan Tidak Ada"
            Combo1 = ""
            Combo1.SetFocus
        End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
        If Not RSPerkiraan.EOF Then
            Text1 = RSPerkiraan!NamaPrk
            Pesan = MsgBox("Yakin akan dihapus..?", vbYesNo)
            If Pesan = vbYes Then
                Dim hapus As String
                hapus = "Delete from Perkiraan where kodeprk='" & Combo1 & "'"
                Conn.Execute hapus
                Call Semula
            Else
                Call Semula
            End If
        Else
            MsgBox "Kode perkiraan tidak ditemukan"
            Combo1.SetFocus
            Exit Sub
        End If
    End If
End If
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Do
Me.Left = Me.Left + 40
Me.Move Me.Left, Me.Top
DoEvents
Loop Until Me.Left > Screen.Width
End Sub

