Attribute VB_Name = "Module2"
#If Win16 Then
  Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
  End Type
#Else
  Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
#End If

#If Win16 Then
  Declare Sub GetWindowRect Lib "User" (ByVal hwnd As _
  Integer, lpRect As RECT)
  Declare Function GetDC Lib "User" (ByVal hwnd As _
  Integer) As Integer
  Declare Function ReleaseDC Lib "User" (ByVal hwnd _
  As Integer, ByVal hdc As Integer) As Integer
  Declare Sub SetBkColor Lib "GDI" (ByVal hdc As _
  Integer, ByVal crColor As Long)
  Declare Sub Rectangle Lib "GDI" (ByVal hdc As _
  Integer, ByVal X1 As Integer, ByVal Y1 As Integer, _
  ByVal X2 As Integer, ByVal Y2 As Integer)
  Declare Function CreateSolidBrush Lib "GDI" (ByVal _
  crColor As Long) As Integer
  Declare Function SelectObject Lib "GDI" (ByVal hdc _
  As Integer, ByVal hObject As Integer) As Integer
  Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
  Declare Function GetWindowRect Lib "user32" (ByVal _
  hwnd As Long, lpRect As RECT) As Long
  Declare Function GetDC Lib "user32" (ByVal hwnd As _
  Long) As Long
  Declare Function ReleaseDC Lib "user32" (ByVal hwnd _
  As Long, ByVal hdc As Long) As Long
  Declare Function SetBkColor Lib "gdi32" (ByVal hdc _
  As Long, ByVal crColor As Long) As Long
  Declare Function Rectangle Lib "gdi32" (ByVal hdc _
  As Long, ByVal X1 As Long, ByVal Y1 As Long, _
  ByVal X2 As Long, ByVal Y2 As Long) As Long
  Declare Function CreateSolidBrush Lib "gdi32" _
  (ByVal crColor As Long) As Long
  Declare Function SelectObject Lib "user32" (ByVal _
  hdc As Long, ByVal hObject As Long) As Long
  Declare Function DeleteObject Lib "gdi32" (ByVal _
  hObject As Long) As Long
#End If

Public Sub ImplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, x%, Y%, Cx%, Cy%
Dim TheScreen As Long
Dim Brush As Long
  GetWindowRect f.hwnd, myRect
  formWidth = (myRect.Right - myRect.Left)
  formHeight = myRect.Bottom - myRect.Top
  TheScreen = GetDC(0)
  Brush = CreateSolidBrush(f.BackColor)
  For i = Movement To 1 Step -1
    Cx = formWidth * (i / Movement)
    Cy = formHeight * (i / Movement)
    x = myRect.Left + (formWidth - Cx) / 2
    Y = myRect.Top + (formHeight - Cy) / 2
    Rectangle TheScreen, x, Y, x + Cx, Y + Cy
  Next i
  x = ReleaseDC(0, TheScreen)
  DeleteObject (Brush)
End Sub







