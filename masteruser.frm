VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form masteruser 
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
      Begin VB.CommandButton cmdkeluaruser 
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton cmdhapususer 
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
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdedituser 
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
         TabIndex        =   4
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdtambahuser 
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
         TabIndex        =   3
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   630
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6615
      Left            =   240
      TabIndex        =   1
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
         Weight          =   400
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
      Caption         =   "MASTER USER"
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
      Width           =   2625
   End
End
Attribute VB_Name = "masteruser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdedituser_Click()
'perintah jika tombol edit diklik
tambahedituser1.Show
tambahedituser1.Luser.Caption = "Edit user1"
tambahedituser1.cmdsimpan.Caption = "&Update"
tambahedituser1.txtnik.Visible = False
tambahedituser1.txtnik1.Visible = True
tambahedituser1.txtnik1.Text = DataGrid1.Columns(0)
tambahedituser1.txtnama.Text = DataGrid1.Columns(1)
tambahedituser1.txtpengguna.Text = DataGrid1.Columns(2)
tambahedituser1.txtpassword.Text = DataGrid1.Columns(3)
tambahedituser1.txtnik.Enabled = False
tambahedituser1.txtpengguna.SetFocus
MDIForm1.Enabled = False
End Sub

Private Sub cmdhapususer_Click()
'perinta hapus data pengguna jika tombol hapus diklik
tanya = MsgBox("Apakah Ingin menghapus " + DataGrid1.Columns(2) + "?", vbQuestion + vbYesNo)
    If tanya = vbYes Then
        Set rspengguna = New ADODB.Recordset
        strsql = "delete from tabel_pengguna where pengguna = '" & DataGrid1.Columns(2) & "'"
        Set rspengguna = con.Execute(strsql, , adCmdText)
        MsgBox "data telah dihapus", vbOKOnly, "Sukses"
        Form_Load
    End If
End Sub

Private Sub cmdkeluaruser_Click()
'perintah jika tombol keluar diklik
Unload Me
MDIForm1.Show
End Sub

Private Sub cmdtambahuser_Click()
'perintah jika tombol tambah diklik
tambahedituser1.Show
tambahedituser1.Luser.Caption = "Tambah user"
tambahedituser1.txtnik.SetFocus
tambahedituser1.cmdsimpan.Caption = "&Simpan"
MDIForm1.Enabled = False
End Sub

Private Sub Form_Load()
'buka koneksi, tabel database dan tampilkan didalam grid
pengguna
End Sub

Private Sub txtcari_Change()
'perintah jika melakukan pencarian dikolom pencarian
bukakoneksi
bukapengguna
Set rspengguna = New ADODB.Recordset
strsql = "SELECT tabel_karyawan.nik, tabel_karyawan.nama, tabel_pengguna.pengguna, tabel_pengguna.pasword " & _
"from tabel_karyawan inner join tabel_pengguna on tabel_karyawan.nik = tabel_pengguna.nik " & _
"where tabel_karyawan.status_user = 'Yes' and tabel_karyawan.status = 'Aktif' and tabel_pengguna.pengguna like '" & txtcari.Text & "%'"
rspengguna.Open strsql, con
    If Not rspengguna.EOF Then 'jika ada, maka tampilkan di datagrid
        gridpengguna
    Else 'jika tidak ada, maka beri peringatan
        MsgBox "Data yang dicari tidak ada", vbOKOnly + vbCritical, "Peringatan"
        Form_Load
    End If
End Sub
