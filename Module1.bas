Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rskaryawan As New ADODB.Recordset
Public rsquerikaryawan As New ADODB.Recordset

Public rsjabatan As New ADODB.Recordset

Public rspengguna As New ADODB.Recordset

Public rsrekapgaji As New ADODB.Recordset

Public rsquerirekapgaji As New ADODB.Recordset
Public strsql As String

    
Public Sub bukakoneksi()
Set con = New ADODB.Connection
con.ConnectionString = "provider=microsoft.jet.oledb.4.0; Data source=" & _
"" & App.Path & "\database\penggajihan.mdb"
con.Open koneksi
con.CursorLocation = adUseClient
End Sub

Public Sub bukakaryawan()
Set rskaryawan = New ADODB.Recordset
strsql = "select tabel_karyawan.nik, tabel_karyawan.nama, tabel_karyawan.alamat, " & _
"tabel_karyawan.nohp, tabel_jabatan.jabatan, tabel_karyawan.status_user, tabel_karyawan.status " & _
"from tabel_karyawan " & _
"inner join tabel_jabatan on tabel_karyawan.id_jabatan = tabel_jabatan.id_jabatan"
rskaryawan.Open strsql, con
End Sub

Public Sub bukajabatan()
Set rsjabatan = New ADODB.Recordset
strsql = "select * from tabel_jabatan"
rsjabatan.Open strsql, con
End Sub
Public Sub bukapengguna()
Set rspengguna = New ADODB.Recordset
strsql = "SELECT tabel_karyawan.nik, tabel_karyawan.nama, tabel_pengguna.pengguna, tabel_pengguna.pasword " & _
"from tabel_karyawan inner join tabel_pengguna on tabel_karyawan.nik = tabel_pengguna.nik " & _
"where tabel_karyawan.status_user = 'Yes' and tabel_karyawan.status = 'Aktif'"
rspengguna.Open strsql, con
End Sub

Public Sub bukarekapgaji()
Set rsrekapgaji = New ADODB.Recordset
strsql = "SELECT tabel_penggajihan.bulan, tabel_penggajihan.tgl, tabel_karyawan.nik, tabel_karyawan.nama, tabel_jabatan.jabatan, tabel_penggajihan.gatot, tabel_penggajihan.totalpotongan, tabel_penggajihan.gajibersih,tabel_penggajihan.status " & _
"FROM (tabel_jabatan INNER JOIN tabel_karyawan ON tabel_jabatan.id_jabatan = tabel_karyawan.id_jabatan) INNER JOIN tabel_penggajihan ON tabel_karyawan.nik = tabel_penggajihan.nik " & _
"where tabel_penggajihan.bulan = '" & transaksigajih.txtbulan.Text & "'"
rsrekapgaji.Open strsql, con
End Sub

Public Sub bukaquerirekapgaji()
Set rsquerirekapgaji = New ADODB.Recordset
strsql = "SELECT tabel_penggajihan.bulan, tabel_penggajihan.tgl, tabel_karyawan.nik, tabel_karyawan.nama, tabel_jabatan.jabatan, tabel_penggajihan.gatot, tabel_penggajihan.totalpotongan, tabel_penggajihan.gajibersih,tabel_penggajihan.status " & _
"FROM (tabel_jabatan INNER JOIN tabel_karyawan ON tabel_jabatan.id_jabatan = tabel_karyawan.id_jabatan) INNER JOIN tabel_penggajihan ON tabel_karyawan.nik = tabel_penggajihan.nik " & _
"where tabel_penggajihan.bulan = '" & transaksigajih.txtbulan.Text & "'"
rsquerirekapgaji.Open strsql, con
End Sub

Public Sub konek()
bukakoneksi
bukajabatan
bukakaryawan
bukapengguna
'bukarekapgaji
End Sub

Public Sub karyawan()
konek
gridkaryawan
End Sub


Public Sub gridkaryawan()
If Not rskaryawan.EOF Then
    Set masterkaryawan.DataGrid2.DataSource = rskaryawan
    With masterkaryawan.DataGrid2
    .Columns(0).Caption = " NIK "
    .Columns(0).Width = 2000
    .Columns(1).Caption = " NAMA "
    .Columns(1).Width = 4000
    .Columns(2).Caption = "ALAMAT"
    .Columns(2).Width = 4000
    .Columns(3).Caption = "NO. TELP"
    .Columns(3).Width = 3000
    .Columns(4).Caption = "JABATAN"
    .Columns(4).Width = 3000
    .Columns(5).Caption = "USER"
    .Columns(5).Width = 2000
    .Columns(6).Caption = "STATUS"
    .Columns(6).Width = 2000
    End With
    masterkaryawan.cmdhapus.Enabled = True
    masterkaryawan.cmdeditkaryawan.Enabled = True
Else
    Set masterkaryawan.DataGrid2.DataSource = rskaryawan
    With masterkaryawan.DataGrid2
    .Columns(0).Caption = " NIK "
    .Columns(0).Width = 2000
    .Columns(1).Caption = " NAMA "
    .Columns(1).Width = 4000
    .Columns(2).Caption = "ALAMAT"
    .Columns(2).Width = 4000
    .Columns(3).Caption = "NO. TELP"
    .Columns(3).Width = 3000
    .Columns(4).Caption = "JABATAN"
    .Columns(4).Width = 3000
    .Columns(5).Caption = "USER"
    .Columns(5).Width = 2000
    .Columns(6).Caption = "STATUS"
    .Columns(6).Width = 2000
    End With
    masterkaryawan.cmdhapus.Enabled = False
    masterkaryawan.cmdeditkaryawan.Enabled = False
End If
End Sub

Public Sub jabatan()
konek
gridjabatan
End Sub


Public Sub gridjabatan()
If Not rsjabatan.EOF Then
    Set masterjabatan.DataGrid1.DataSource = rsjabatan
    With masterjabatan.DataGrid1
    .Columns(0).Caption = " ID JABATAN "
    .Columns(0).Width = 5000
    .Columns(1).Caption = " NAMA JABATAN "
    .Columns(1).Width = 15000
    End With
    masterjabatan.cmdeditjabatan.Enabled = True
    masterjabatan.cmdhapusjabatan.Enabled = True
Else
    Set masterjabatan.DataGrid1.DataSource = rsjabatan
    With masterjabatan.DataGrid1
    .Columns(0).Caption = " ID JABATAN "
    .Columns(0).Width = 5000
    .Columns(1).Caption = " NAMA JABATAN "
    .Columns(1).Width = 15000
    End With
    masterjabatan.cmdeditjabatan.Enabled = False
    masterjabatan.cmdhapusjabatan.Enabled = False
End If
End Sub

Public Sub pengguna()
konek
gridpengguna
End Sub

Public Sub gridpengguna()
If Not rspengguna.EOF Then
    Set masteruser.DataGrid1.DataSource = rspengguna
    With masteruser.DataGrid1
    .Columns(0).Caption = " NIK "
    .Columns(0).Width = 4500
    .Columns(1).Caption = " NAMA "
    .Columns(1).Width = 7000
    .Columns(2).Caption = "PENGGUNA"
    .Columns(2).Width = 4000
    .Columns(3).Caption = "PASSWORD"
    .Columns(3).Width = 4000
    End With
    masteruser.cmdedituser.Enabled = True
    masteruser.cmdhapususer.Enabled = True
Else
    Set masteruser.DataGrid1.DataSource = rspengguna
    With masteruser.DataGrid1
    .Columns(0).Caption = " NIK "
    .Columns(0).Width = 4500
    .Columns(1).Caption = " NAMA "
    .Columns(1).Width = 7000
    .Columns(2).Caption = "PENGGUNA"
    .Columns(2).Width = 4000
    .Columns(3).Caption = "PASSWORD"
    .Columns(3).Width = 4000
    End With
    masteruser.cmdedituser.Enabled = False
    masteruser.cmdhapususer.Enabled = False
End If
End Sub

Public Sub rekapgaji()
konek
gridquerirekapgaji
End Sub

Public Sub gridquerirekapgaji()
If Not rsquerirekapgaji.EOF Then
    Set transaksigajih.DataGrid1.DataSource = rsquerirekapgaji
    With transaksigajih.DataGrid1
    .Columns(0).Caption = " PERIODE "
    .Columns(0).Width = 1500
    .Columns(1).Caption = " TANGGAL "
    .Columns(1).Width = 1500
    .Columns(2).Caption = "NIK"
    .Columns(2).Width = 1500
    .Columns(3).Caption = "NAMA"
    .Columns(3).Width = 3000
    .Columns(4).Caption = "JABATAN"
    .Columns(4).Width = 3000
    .Columns(5).Caption = "GAJI TOTAL"
    .Columns(5).Width = 2500
    .Columns(6).Caption = "TOTAL POTONGAN"
    .Columns(6).Width = 2500
    .Columns(7).Caption = "GAJI BERSIH"
    .Columns(7).Width = 2500
    .Columns(8).Caption = "STATUS"
    .Columns(8).Width = 1500
    End With
    transaksigajih.cmddetailgajih.Caption = "&Detail   [F1]"
    transaksigajih.cmdkonfirmasi.Visible = True
Else
    Set transaksigajih.DataGrid1.DataSource = rsquerirekapgaji
    With transaksigajih.DataGrid1
    .Columns(0).Caption = " PERIODE "
    .Columns(0).Width = 1500
    .Columns(1).Caption = " TANGGAL "
    .Columns(1).Width = 1500
    .Columns(2).Caption = "NIK"
    .Columns(2).Width = 1500
    .Columns(3).Caption = "NAMA"
    .Columns(3).Width = 3000
    .Columns(4).Caption = "JABATAN"
    .Columns(4).Width = 3000
    .Columns(5).Caption = "GAJI TOTAL"
    .Columns(5).Width = 2500
    .Columns(6).Caption = "TOTAL POTONGAN"
    .Columns(6).Width = 2500
    .Columns(7).Caption = "GAJI BERSIH"
    .Columns(7).Width = 2500
    .Columns(8).Caption = "STATUS"
    .Columns(8).Width = 1500
    End With
     transaksigajih.cmddetailgajih.Caption = "&Rekap   [F1]"
    transaksigajih.cmdkonfirmasi.Visible = False
End If
End Sub
