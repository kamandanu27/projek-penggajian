VERSION 5.00
Begin VB.Form tambahedituser1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnik1 
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtnama 
      Enabled         =   0   'False
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2160
      Width           =   4215
   End
   Begin VB.ComboBox txtnik 
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
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtpassword 
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtpengguna 
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
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   7680
      Width           =   8055
      Begin VB.CommandButton Command2 
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
         TabIndex        =   12
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
         TabIndex        =   11
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
      Begin VB.Label Luser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tambah User"
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
         Width           =   2190
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA"
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
      TabIndex        =   9
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIK"
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
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      TabIndex        =   4
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA PENGGUNA"
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
      TabIndex        =   3
      Top             =   2880
      Width           =   2355
   End
End
Attribute VB_Name = "tambahedituser1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsimpan_Click()
If cmdsimpan.Caption = "&Update" Then

    If txtpengguna.Text = "" Then
    MsgBox "Nama Pengguna Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtpengguna.SetFocus
    Exit Sub
    End If

    If txtpassword.Text = "" Then
    MsgBox "Password Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtpassword.SetFocus
    Exit Sub
    End If

update
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masteruser.Refresh
bukakoneksi
bukapengguna
gridpengguna
End If

If cmdsimpan.Caption = "&Simpan" Then
    If txtpengguna.Text = "" Then
    MsgBox "Nama Pengguna Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtpengguna.SetFocus
    Exit Sub
    End If
    
    If txtpassword.Text = "" Then
    MsgBox "Password Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtpassword.SetFocus
    Exit Sub
    End If
    
simpan
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masteruser.Refresh
bukakoneksi
bukapengguna
gridpengguna
End If

End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub Form_Load()
konek
combonik
End Sub

Private Sub update()
Set rspengguna = New ADODB.Recordset
strsql = "update tabel_pengguna set pengguna = '" & txtpengguna.Text & "', pasword = '" & txtpassword.Text & "' " & _
        "where pengguna = '" & txtpengguna.Text & "'"
        Set rspengguna = con.Execute(strsql, , adCmdText)
        MsgBox "Data Telah Diperbarui", vbOKOnly, "SUKSES"
End Sub

Private Sub simpan()
Set rspengguna = New ADODB.Recordset
strsql = "insert into tabel_pengguna values" & _
    "('" & Split(txtnik.Text, "|")(0) & " '," & _
    "'" & Trim(txtpengguna.Text) & " '," & _
      "'" & Trim(txtpassword.Text) & "')"
    rspengguna.Open strsql, con, adOpenDynamic, adLockOptimistic
    MsgBox "data telah disimpan", vbOKOnly, "SUKSES"
End Sub

Sub combonik()
Set rspengguna = New ADODB.Recordset
strsql = "SELECT tabel_karyawan.nik as nik, tabel_karyawan.nama as nama from tabel_karyawan " & _
"where status_user = 'Yes' and status = 'Aktif'"
rspengguna.Open strsql, con, adOpenDynamic, adLockOptimistic
rspengguna.Requery
With rspengguna
If .EOF And .BOF Then
txtpengguna.Text = ""
Else
txtnik.Clear
Do Until .EOF
txtnik.AddItem ![nik] _
+ " | " + ![nama]
.MoveNext
Loop
.MoveFirst
End If
End With
End Sub

'Private Sub txtnik_Change()
''rspengguna.Filter = "nama='" & Split(txtnik.Text, "|")(0) & "'"
''    If Not rspengguna.EOF Then
'        txtnama.Text = txtnik.Text
'    'End If
'End Sub

Private Sub txtnik_Click()
rspengguna.Filter = "nik='" & Split(txtnik.Text, "|")(0) & "'"
    If Not rspengguna.EOF Then
        txtnama.Text = rspengguna!nama
    End If
End Sub

Private Sub txtpengguna_GotFocus()
txtpengguna.SelStart = 0
txtpengguna.SelLength = Len(txtpengguna)
End Sub

Private Sub txtpassword_GotFocus()
txtpassword.SelStart = 0
txtpassword.SelLength = Len(txtpassword)
End Sub
