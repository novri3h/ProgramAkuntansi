VERSION 5.00
Begin VB.Form GantiPass 
   Caption         =   "Ganti Password"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3825
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
   ScaleHeight     =   2370
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Konfirmasi Pwd Baru"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password Baru"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password Lama"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "GantiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSKASIR.Open "select * from kasir where namaksr='" & Text1 & "'", Conn
    If Not RSKASIR.EOF Then
        Text2.SetFocus
    Else
        MsgBox "nama kasir tidak terdaftar"
        Text1.SetFocus
        Text1 = ""
    End If
End If

End Sub

Private Sub Text2_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSKASIR.Open "select * from kasir where namaksr='" & Text1 & "' and passwordksr='" & Text2 & "'", Conn
    If Not RSKASIR.EOF Then
        Text3.SetFocus
    Else
        MsgBox "password salah "
        Text2.SetFocus
        Text2 = ""
    End If
End If

End Sub

Private Sub Text3_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Text3 = "" Then
        MsgBox "password baru belum dibuat"
        Text3.SetFocus
    Else
        Text4.SetFocus
    End If
End If
End Sub

Private Sub Text4_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Text4 <> Text3 Then
        MsgBox "password konfirmasi tidak sama"
        Text4.SetFocus
        Text4 = ""
    Else
        Pesan = MsgBox("yakin password akan diganti", vbYesNo)
        If Pesan = vbYes Then
            Dim editpwd As String
            editpwd = "update kasir set passwordksr='" & Text4 & "' where namaksr='" & Text1 & "' and passwordksr='" & Text2 & "'"
            Conn.Execute editpwd
            Unload Me
        Else
            Unload Me
        End If
    End If
End If

End Sub

