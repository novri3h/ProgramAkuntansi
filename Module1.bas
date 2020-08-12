Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSPerkiraan As ADODB.Recordset
Public RSPemesan As ADODB.Recordset
Public RSMasterPO As ADODB.Recordset
Public RSDetailPO As ADODB.Recordset
Public RSBukuBesar As ADODB.Recordset
Public RSKASIR As ADODB.Recordset
Public RSKas As ADODB.Recordset
Public RSArusKas As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSPerkiraan = New ADODB.Recordset
Set RSPemesan = New ADODB.Recordset
Set RSMasterPO = New ADODB.Recordset
Set RSDetailPO = New ADODB.Recordset
Set RSBukuBesar = New ADODB.Recordset
Set RSKASIR = New ADODB.Recordset
Set RSKas = New ADODB.Recordset
Set RSArusKas = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBAKN.mdb"
End Sub


