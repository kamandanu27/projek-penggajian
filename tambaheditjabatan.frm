VERSION 5.00
Begin VB.Form tambaheditjabatan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtjabatan 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox txtidjabatan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   8055
      Begin VB.CommandButton cmdbatal 
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Label Ljabatan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tambah Jabatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   435
         Left            =   2280
         TabIndex        =   2
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA JABATAN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID JABATAN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   1545
   End
End
Attribute VB_Name = "tambaheditjabatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdbatal_Click()
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub cmdsimpan_Click()
'jika caption tombol Update
If cmdsimpan.Caption = "&Update" Then
    If txtidjabatan.Text = "" Then
    MsgBox "ID Jabatan Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnik.SetFocus
    Exit Sub
    End If
    
    
    If txtjabatan.Text = "" Then
    MsgBox "Jabatan Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnama.SetFocus
    Exit Sub
    End If
    
update
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masterjabatan.Refresh
jabatan
End If

'jika caption tombol Simpan
If cmdsimpan.Caption = "&Simpan" Then
    If txtidjabatan.Text = "" Then
    MsgBox "ID Jabatan Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnik.SetFocus
    Exit Sub
    End If
    
    
    If txtjabatan.Text = "" Then
    MsgBox "Nama Jabatan Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnama.SetFocus
    Exit Sub
    End If

simpan
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masterjabatan.Refresh
jabatan
End If

End Sub

Private Sub update()
Set rsjabatan = New ADODB.Recordset
strsql = " update tabel_jabatan set jabatan = '" & txtjabatan.Text & "' where id_jabatan = '" & txtidjabatan.Text & "'"
        Set rsjabatan = con.Execute(strsql, , adCmdText)
        MsgBox "Data Berhasil Diperbarui", vbOKOnly, "SUKSES"
End Sub

Private Sub simpan()
Set rsjabatan = New ADODB.Recordset
strsql = "insert into tabel_jabatan values" & _
    "('" & txtidjabatan.Text & "'," & _
    "'" & Trim(txtjabatan.Text) & "')"
    rsjabatan.Open strsql, con, adOpenDynamic, adLockOptimistic
    MsgBox "Data Telah Disimpan", vbOKOnly, "SUKSES"
End Sub

Private Sub Form_Load()
bukakoneksi
bukakaryawan
End Sub

Private Sub txtjabatan_GotFocus()
txtjabatan.SelStart = 0
txtjabatan.SelLength = Len(txtjabatan)
End Sub
