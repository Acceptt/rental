Attribute VB_Name = "Module1"
Public Db As Connection

Public Function Aktif_Koneksi() As Boolean
Set Db = New ADODB.Connection
Db.CursorLocation = adUseClient
Db.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\database.accdb;"
End Function
