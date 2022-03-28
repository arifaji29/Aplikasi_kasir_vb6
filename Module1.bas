Attribute VB_Name = "Module1"
Public Koneksi As New ADODB.Connection
Public RSKasir As ADODB.Recordset
Public RSBarang As ADODB.Recordset
Public RSPelanggan As ADODB.Recordset
Public RSJual As ADODB.Recordset

Public Sub bukaDB()
Set Koneksi = New ADODB.Connection
Set RSKasir = New ADODB.Recordset
Set RSBarang = New ADODB.Recordset
Set RSPelanggan = New ADODB.Recordset
Set RSJual = New ADODB.Recordset
Koneksi.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Konek_DBKasir"
End Sub



