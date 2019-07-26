VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form transaksigajih 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cetakslipgaji 
      Left            =   7080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtbantu 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Text            =   "0"
      Top             =   8520
      Width           =   375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   16200
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox txtbulan 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   7440
      Width           =   19935
      Begin VB.CommandButton cmdkonfirmasi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton tblkeluaruser 
         Caption         =   "Keluar   [ F5 ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   17760
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmddetailgajih 
         Caption         =   "Detail   [F1]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   11033
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA TRANSAKSI PENGGAJIAN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5730
   End
End
Attribute VB_Name = "transaksigajih"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub konfirmasi()
If DataGrid1.Columns(5) < 1 Then
MsgBox "Konfirmasi Tidak Dapat Diproses! Gapok Tidak Boleh Kosong", vbOKOnly, "Peringatan"
Exit Sub
Else
Set rsrekapgaji = New ADODB.Recordset
strsql = " update tabel_penggajihan set status = 'Selesai' " & _
        "where nik = '" & DataGrid1.Columns(2) & "' and bulan = '" & txtbulan.Text & "'"
        Set rsrekapgaji = con.Execute(strsql, , adCmdText)
    bukakoneksi
    bukaquerirekapgaji
    gridquerirekapgaji
'    cmdkonfirmasi.Caption = "&Cetak Slip Gaji   [F2]"
        MsgBox "Data Telah Diperbarui", vbOKOnly, "SUKSES"
        
Form_Load
End If
End Sub

Private Sub rekap()
Set rsquerikaryawan = New ADODB.Recordset
strsql = "select * from tabel_karyawan where status = 'Aktif'"
rsquerikaryawan.Open strsql, con

While Not rsquerikaryawan.EOF
Dim nama As String
nama = rsquerikaryawan!nik
        
    Set rsquerirekapgaji = New ADODB.Recordset
    strsql = "insert into tabel_penggajihan values" & _
    "('" & txtbulan.Text & " '," & _
    "'" & MDIForm1.StatusBar1.Panels(4) & "'," & _
     "'" & Trim(nama) & "'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'0'," & _
      "'Proses')"
    rsquerirekapgaji.Open strsql, con, adOpenDynamic, adLockOptimistic
    
    rsquerikaryawan.MoveNext
Wend
    MsgBox "data telah disimpan", vbOKOnly, "SUKSES"
    Form_Load
End Sub

Private Sub cetak_slipgaji()
With cetakslipgaji
.Reset
.SelectionFormula = "{tabel_penggajihan.nik}='" & DataGrid1.Columns(2) & "'"
.ReportFileName = "" & App.Path & "\report\slipgaji.rpt"
.Destination = crptToWindow
.WindowState = crptMaximized
.DiscardSavedData = True
.Action = 1
End With
End Sub
Private Sub cmddetailgajih_Click()
If transaksigajih.cmddetailgajih.Caption = "&Detail   [F1]" Then
prosespenggajian.Show
prosespenggajian.txtnik.Text = DataGrid1.Columns(2)
prosespenggajian.txtnama.Text = DataGrid1.Columns(3)
prosespenggajian.txtjabatan.Text = DataGrid1.Columns(4)

prosespenggajian.txtbulan.Text = DataGrid1.Columns(0)
prosespenggajian.txttgl.Text = DataGrid1.Columns(1)
prosespenggajian.txtstatus.Text = DataGrid1.Columns(8)
MDIForm1.Enabled = False
Else
tanya = MsgBox("Apakah Ingin Merekap Karyawan Aktif?", vbQuestion + vbYesNo)
    If tanya = vbYes Then
    rekap
    End If
End If
End Sub

Private Sub cmdkonfirmasi_Click()
If transaksigajih.cmdkonfirmasi.Caption = "&Konfirmasi   [F2]" Then
tanya = MsgBox("Apakah Ingin Memproses Gaji " + DataGrid1.Columns(3) + " ?", vbQuestion + vbYesNo)
    If tanya = vbYes Then
    konfirmasi
    End If
Else
tanya = MsgBox("Apakah Ingin Mencetak Slip Gaji " + DataGrid1.Columns(3) + " ?", vbQuestion + vbYesNo)
    If tanya = vbYes Then
    cetak_slipgaji
    End If
End If
End Sub

Private Sub DataGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not rsquerirekapgaji.EOF Then
    If DataGrid1.Columns(8).Text = "Proses" Then
    cmdkonfirmasi.Caption = "&Konfirmasi   [F2]"
    Else
    cmdkonfirmasi.Caption = "&Cetak Slip Gaji   [F2]"
    End If
End If
End Sub

Private Sub Form_Load()
bukakoneksi
bukaquerirekapgaji
rekapgaji

txtbulan.AddItem ("Januari 2019")
txtbulan.AddItem ("Februari 2019")
txtbulan.AddItem ("Maret 2019")
txtbulan.AddItem ("April 2019")
txtbulan.AddItem ("Mei 2019")
txtbulan.AddItem ("Juni 2019")
End Sub

Private Sub tbldetailgajih_Click()
prosespenggajian.Show
MDIForm1.Enabled = False
End Sub

Private Sub tbledituser_Click()
tambahedituser1.Show
tambahedituser1.Luser.Caption = "Edit user"
MDIForm1.Enabled = False
End Sub


Private Sub tblkeluaruser_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub tbltambahuser_Click()
tambahedituser1.Show
tambahedituser1.Luser.Caption = "Tambah user"
MDIForm1.Enabled = False
End Sub

Private Sub txtbulan_Change()
bukakoneksi
bukaquerirekapgaji
gridquerirekapgaji
End Sub

Private Sub txtbulan_Click()
bukakoneksi
bukaquerirekapgaji
gridquerirekapgaji
End Sub

