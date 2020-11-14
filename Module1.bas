Attribute VB_Name = "Koneksi"
Option Explicit
Public Conn As ADODB.Connection
Public Rs As ADODB.Recordset
Dim AA As String
Public Sub BukaData()
On Error GoTo cek
Set Conn = New ADODB.Connection
Conn.CursorLocation = adUseClient
'AA = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Fossil.mdb;Persist Security Info=False"
'Conn.Open AA
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Fossil.mdb;Persist Security Info=False"
 Exit Sub

cek:
MsgBox "Gagal menghubungkan ke Database ! Kesalahan pada : " & Err.Description, vbCritical, "Peringatan"

End Sub
