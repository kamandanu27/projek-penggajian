VERSION 5.00
Begin VB.Form tambaheditkaryawan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtstatus 
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
      ItemData        =   "tambaheditkaryawan.frx":0000
      Left            =   2280
      List            =   "tambaheditkaryawan.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox txtuser 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox txtjabatan 
      CausesValidation=   0   'False
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
      ItemData        =   "tambaheditkaryawan.frx":0004
      Left            =   2280
      List            =   "tambaheditkaryawan.frx":0006
      TabIndex        =   8
      Text            =   "txtjabatan"
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtnotlp 
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
      Left            =   2280
      MaxLength       =   13
      TabIndex        =   7
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox txtalamat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2280
      MaxLength       =   70
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox txtnama 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox txtnik 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   8040
      Width           =   8055
      Begin VB.CommandButton cmdbatal 
         Caption         =   "&Batal"
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
         TabIndex        =   14
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
         TabIndex        =   12
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
      Begin VB.Label Lkaryawan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tambah Karyawan"
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
         Width           =   3000
      End
   End
   Begin VB.TextBox txtjabatan1 
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
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
      TabIndex        =   18
      Top             =   5880
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER"
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
      TabIndex        =   17
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JABATAN"
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
      TabIndex        =   16
      Top             =   4440
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO. TELP"
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
      TabIndex        =   15
      Top             =   3720
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT"
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
      TabIndex        =   13
      Top             =   2760
      Width           =   1110
   End
   Begin VB.Label Label3 
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
      TabIndex        =   11
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label Label2 
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
      TabIndex        =   4
      Top             =   1320
      Width           =   480
   End
End
Attribute VB_Name = "tambaheditkaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsimpan_Click()
If cmdsimpan.Caption = "&Update" Then
    If txtnik.Text = "" Then
    MsgBox "NIK Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnik.SetFocus
    Exit Sub
    End If
    
    
    If txtnama.Text = "" Then
    MsgBox "Nama Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnama.SetFocus
    Exit Sub
    End If
    
    If txtnotlp.Text = "" Then
    MsgBox "No Telepon Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnotlp.SetFocus
    Exit Sub
    End If
    
    If txtalamat.Text = "" Then
    MsgBox "Alamat Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtalamat.SetFocus
    Exit Sub
    End If
    
    If txtnotlp.Text = "" Then
    MsgBox "No Tlp Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnotlp.SetFocus
    Exit Sub
    End If
    
update
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masterkaryawan.Refresh
karyawan
End If

If cmdsimpan.Caption = "&Simpan" Then
    If txtnik.Text = "" Then
    MsgBox "NIK Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnik.SetFocus
    Exit Sub
    End If
    
    
    If txtnama.Text = "" Then
    MsgBox "Nama Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnama.SetFocus
    Exit Sub
    End If
    
    If txtnotlp.Text = "" Then
    MsgBox "No Telepon Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnotlp.SetFocus
    Exit Sub
    End If
    
    If txtalamat.Text = "" Then
    MsgBox "Alamat Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtalamat.SetFocus
    Exit Sub
    End If
    
    If txtnotlp.Text = "" Then
    MsgBox "No Tlp Tidak Boleh Kosong", vbOKOnly + vbCritical, "Peringatan"
    txtnotlp.SetFocus
    Exit Sub
    End If

simpan
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
masterkaryawan.Refresh
karyawan
End If

End Sub

Private Sub cmdbatal_Click()
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub Form_Load()
bukakoneksi
bukakaryawan

combojabatan

txtuser.AddItem ("Yes")
txtuser.AddItem ("No")

txtstatus.AddItem ("Aktif")
txtstatus.AddItem ("Resign")
End Sub



Private Sub update()
Set rskaryawan = New ADODB.Recordset
strsql = " update tabel_karyawan set nama = '" & txtnama.Text & "' " & _
        ",alamat = '" & txtalamat.Text & "' " & _
        ",nohp = '" & txtnotlp.Text & "' " & _
        ",id_jabatan = '" & txtjabatan1.Text & "' " & _
        ",status_user = '" & txtuser.Text & "' " & _
        ",status = '" & txtstatus.Text & "' where nik = '" & txtnik.Text & "'"
        Set rskaryawan = con.Execute(strsql, , adCmdText)
        MsgBox ("Data Telah Diperbarui")
End Sub

Private Sub simpan()
Set rskaryawan = New ADODB.Recordset
strsql = "insert into tabel_karyawan values" & _
    "('" & txtnik.Text & " '," & _
    "'" & Trim(txtnama.Text) & " '," & _
     "'" & Trim(txtalamat.Text) & "'," & _
     "'" & txtnotlp.Text & "'," & _
     "'" & Trim(txtjabatan1.Text) & "'," & _
     "'" & Trim(txtuser.Text) & "'," & _
      "'" & Trim(txtstatus.Text) & "')"
    rskaryawan.Open strsql, con, adOpenDynamic, adLockOptimistic
    MsgBox ("data telah disimpan")
End Sub

Sub combojabatan()
Set rsjabatan = New ADODB.Recordset
strsql = "select * from tabel_jabatan"
rsjabatan.Open strsql, con, adOpenDynamic, adLockOptimistic
rsjabatan.Requery
With rsjabatan
If .EOF And .BOF Then
txtjabatan.Text = ""
Else
txtjabatan.Clear
Do Until .EOF
txtjabatan.AddItem ![id_jabatan] _
+ " | " + ![jabatan]
.MoveNext
Loop
.MoveFirst
End If
End With
End Sub

Private Sub cekkosong()
If txtnik.Text = "" Then
MsgBox ("NIK Tidak Boleh Kosong")
txtnik.SetFocus
Exit Sub
End If


If txtnama.Text = "" Then
MsgBox ("Nama Tidak Boleh Kosong")
txtnama.SetFocus
Exit Sub
End If

If txtnotlp.Text = "" Then
MsgBox ("No Telepon Tidak Boleh Kosong")
txtnotlp.SetFocus
Exit Sub
End If

If txtalamat.Text = "" Then
MsgBox ("Alamat Tidak Boleh Kosong")
txtalamat.SetFocus
Exit Sub
End If

If txtnotlp.Text = "" Then
MsgBox ("No Tlp Tidak Boleh Kosong")
txtnotlp.SetFocus
Exit Sub
End If


End Sub

Private Sub txtjabatan_Change()
'txtjabatan1.Text = txtjabatan.Text
rsjabatan.Filter = "jabatan='" & Split(txtjabatan.Text, "|")(0) & "'"
    If Not rsjabatan.EOF Then
        txtjabatan1.Text = rsjabatan!id_jabatan
    End If
End Sub

Private Sub txtjabatan_Click()
rsjabatan.Filter = "id_jabatan='" & Left(txtjabatan.Text, 3) & "'"
    If Not rsjabatan.EOF Then
        txtjabatan1.Text = rsjabatan!id_jabatan
    End If
End Sub


Private Sub txtnama_GotFocus()
txtnama.SelStart = 0
txtnama.SelLength = Len(txtnama)
End Sub

Private Sub txtalamat_GotFocus()
txtalamat.SelStart = 0
txtalamat.SelLength = Len(txtalamat)
End Sub

Private Sub txtnotlp_GotFocus()
txtnotlp.SelStart = 0
txtnotlp.SelLength = Len(txtnotlp)
End Sub

Private Sub txtnotlp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{TAB}"
KeyAscii = 0
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
Beep
KeyAscii = 0
End If
End Sub
