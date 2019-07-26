VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdmasuk 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtpasword 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox txtpengguna 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   0
      Width           =   5775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   555
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   1500
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbatal_Click()
Unload Me 'keluar
End Sub

Private Sub cmdmasuk_Click()
'cek user,password,status user dan status karyawan di database
On Error Resume Next
Set rspengguna = New ADODB.Recordset
strsql = "SELECT * " & _
"from tabel_karyawan inner join tabel_pengguna on tabel_karyawan.nik = tabel_pengguna.nik " & _
"where tabel_karyawan.status_user = 'Yes' and tabel_karyawan.status = 'Aktif' and " & _
"tabel_pengguna.pengguna = '" & txtpengguna.Text & "' and tabel_pengguna.pasword = '" & txtpasword.Text & "'"
rspengguna.Open strsql, con, adOpenDynamic, adLockBatchOptimistic, adCmdText

    'jika user dan password benar, maka masuk ke menu utama
    If Not rspengguna.EOF Then
        MDIForm1.Show
        Unload Me
    
    'jika user dan password salah, maka hapus user dan password
    Else
        ya = MsgBox("Username / Password Salah ", vbInformation + vbOKOnly, "KONFIRMASI")
            If ya = vbOK Then
                txtpengguna.Text = ""
                txtpasword.Text = ""
                txtpengguna.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub Form_Load()
'buka koneksi dan buka tabel database
bukakoneksi
bukapengguna
End Sub
