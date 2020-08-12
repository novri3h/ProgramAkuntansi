VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   1575
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKodeKsr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   2000
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtPasswordKsr 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtNamaKsr 
         Height          =   350
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   2000
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Password"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nama"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1005
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A As Byte
Dim B As Byte

Private Sub Form_Load()
TxtNamaKsr.MaxLength = 30
TxtPasswordKsr.MaxLength = 10
TxtPasswordKsr.PasswordChar = "X"
TxtPasswordKsr.Enabled = False
TxtKodeKsr.Enabled = False
End Sub

Private Sub TxtNamaKsr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    RSKASIR.Open "Select NamaKsr from Kasir where NamaKsr ='" & TxtNamaKsr & "'", Conn
    If RSKASIR.EOF Then
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal"
            TxtNamaKsr = ""
            TxtNamaKsr.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & TxtNamaKsr & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            Unload Me
        End If
    Else
        TxtNamaKsr.Enabled = False
        TxtPasswordKsr.Enabled = True
        TxtPasswordKsr.SetFocus
    End If
End If
End Sub


Private Sub txtpasswordksr_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
Dim KodeKasir As String
Dim NamaKasir As String
If Keyascii = 13 Then
    Call BukaDB
    RSKASIR.Open "Select * from Kasir where NamaKsr ='" & TxtNamaKsr & "' and PasswordKsr='" & TxtPasswordKsr & "'", Conn
    If RSKASIR.EOF Then
        B = B + 1
        If 1 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 2 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            TxtPasswordKsr = ""
            TxtPasswordKsr.SetFocus
        ElseIf 3 - B = 0 Then
            MsgBox "Kesempatan ke " & B & " Salah"
            Unload Me
        End If
    Else
        Unload Me
        Menu.Show
        Order.Caption = "Data Pesanan " & TxtNamaKsr
    End If
End If
End Sub

