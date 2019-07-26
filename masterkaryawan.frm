VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form masterkaryawan 
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   6615
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   11668
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   7440
      Width           =   19935
      Begin VB.CommandButton tblkeluarkaryawan 
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
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtcari 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9000
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "Hapus   [ F3 ]"
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
         Left            =   4320
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdeditkaryawan 
         Caption         =   "Edit   [ F2 ]"
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
         Left            =   2400
         MaskColor       =   &H00C00000&
         TabIndex        =   3
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton tbltambahkaryawan 
         Caption         =   "Tambah   [F1]"
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
         TabIndex        =   2
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   8400
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[ F4 ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   12720
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER KARYAWAN"
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
      Top             =   120
      Width           =   3885
   End
End
Attribute VB_Name = "masterkaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'panggil koneksi, tabel database dan menampilkanya didalam datagrid
karyawan
End Sub

Private Sub cmdeditkaryawan_Click()
'perintah jika tombol edit diklik
tambaheditkaryawan.Show
tambaheditkaryawan.Lkaryawan.Caption = "Edit Karyawan"
tambaheditkaryawan.cmdsimpan.Caption = "&Update"
tambaheditkaryawan.txtnik.Text = DataGrid2.Columns(0)
tambaheditkaryawan.txtnama.Text = DataGrid2.Columns(1)
tambaheditkaryawan.txtalamat.Text = DataGrid2.Columns(2)
tambaheditkaryawan.txtnotlp.Text = DataGrid2.Columns(3)
tambaheditkaryawan.txtjabatan.Text = DataGrid2.Columns(4)
tambaheditkaryawan.txtuser.Text = DataGrid2.Columns(5)
tambaheditkaryawan.txtstatus.Text = DataGrid2.Columns(6)
tambaheditkaryawan.txtnik.Enabled = False
tambaheditkaryawan.txtnama.SetFocus
MDIForm1.Enabled = False
End Sub

Private Sub cmdhapus_Click()
'perintah menghapus data karyawan jika tombol hapus diklik
tanya = MsgBox("Apakah Ingin menghapus " + DataGrid2.Columns(1) + "?", vbQuestion + vbYesNo)
    If tanya = vbYes Then
        Set rskaryawan = New ADODB.Recordset
        strsql = "delete from tabel_karyawan where NIK = '" & DataGrid2.Columns(0) & "'"
        Set rskaryawan = con.Execute(strsql, , adCmdText)
        MsgBox "data telah dihapus", vbOKOnly, "Sukses"
        Form_Load
    End If
End Sub

Private Sub tblkeluarkaryawan_Click()
'perintah jika tombol keluar diklik
Unload Me
MDIForm1.Show
End Sub

Private Sub tbltambahkaryawan_Click()
'perintah jika tombol tambah diklik
tambaheditkaryawan.Show
tambaheditkaryawan.Lkaryawan.Caption = "Tambah Karyawan"
tambaheditkaryawan.txtnik.SetFocus
tambaheditkaryawan.cmdsimpan.Caption = "&Simpan"
MDIForm1.Enabled = False
End Sub


Private Sub txtcari_Change()
'perintah ketika melakukan pencarian dikolom cari
bukakoneksi
bukakaryawan
Set rskaryawan = New ADODB.Recordset
strsql = "select * from tabel_karyawan where nama like '" & txtcari & "%'"
rskaryawan.Open strsql, con
    If Not rskaryawan.EOF Then 'jika ada, maka tampilka didalam grid
        gridkaryawan
    Else 'jika tidak ada maka beri peringatan
        MsgBox "Data yang dicari tidak ada", vbOKOnly + vbCritical, "Peringatan"
        Form_Load
    End If
End Sub
