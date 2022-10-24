Attribute VB_Name = "Module1"
'mendefinisikan objek'
Public KON As New ADODB.Connection
Public rspelayanan As New ADODB.Recordset
Public rsuser As New ADODB.Recordset
Public rstransaksi As New ADODB.Recordset
Public rsmember As New ADODB.Recordset
Public rstemp As New ADODB.Recordset
Public rsdetail As New ADODB.Recordset


Sub koneksi()
'membuka koneksi'
Set KON = New ADODB.Connection
Set rspelayanan = New ADODB.Recordset
Set rsuser = New ADODB.Recordset
Set rstransaksi = New ADODB.Recordset
Set rsmember = New ADODB.Recordset
Set rstemp = New ADODB.Recordset
Set rsdetail = New ADODB.Recordset
KON.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=laundry;"
KON.Open
End Sub
