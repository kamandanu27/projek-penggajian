VERSION 5.00
Begin VB.Form prosespenggajian 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROSES PENGGAJIHAN"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttotalbersih 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   43
      Text            =   "RP. 10,000,000"
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox txttgl 
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
      Height          =   375
      Left            =   9600
      TabIndex        =   37
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   615
      Left            =   480
      TabIndex        =   34
      Top             =   7080
      Width           =   11775
      Begin VB.TextBox txtabsen 
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
         Left            =   2760
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3600
         TabIndex        =   36
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Kehadiran"
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
         Left            =   480
         TabIndex        =   35
         Top             =   120
         Width           =   1980
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3375
      Left            =   6600
      TabIndex        =   26
      Top             =   3600
      Width           =   5655
      Begin VB.TextBox txttotalpotongan 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   40
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtkehadiran 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtpph 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtkerja 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtkesehatan 
         Alignment       =   1  'Right Justify
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
         Left            =   2400
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Potongan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   39
         Top             =   2760
         Width           =   1845
      End
      Begin VB.Line Line3 
         X1              =   360
         X2              =   5040
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kehadiran"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   33
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   32
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tenaga Kerja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kesehatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Potongan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   4920
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3375
      Left            =   600
      TabIndex        =   24
      Top             =   3600
      Width           =   5775
      Begin VB.TextBox txttotalpendapatan 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         TabIndex        =   42
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txttunjangan 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         TabIndex        =   1
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtgapok 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         TabIndex        =   0
         Top             =   600
         Width           =   2535
      End
      Begin VB.Line Line4 
         X1              =   360
         X2              =   5160
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pendapatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   41
         Top             =   2760
         Width           =   2145
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tunjangan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   29
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gapok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   28
         Top             =   600
         Width           =   810
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5160
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendapatan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtbulan 
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
      Height          =   375
      Left            =   9960
      TabIndex        =   22
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtalamat 
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
      Height          =   795
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Width           =   5295
   End
   Begin VB.TextBox txtjabatan 
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
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   2880
      Width           =   3495
   End
   Begin VB.TextBox txtstatus 
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
      Height          =   375
      Left            =   10320
      TabIndex        =   18
      Top             =   960
      Width           =   1935
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
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox txtnik 
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
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   8160
      Width           =   12855
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
         Left            =   10920
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "&Simpan"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12795
      TabIndex        =   8
      Top             =   0
      Width           =   12855
      Begin VB.Label Luser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses Penggajian"
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
         Left            =   4680
         TabIndex        =   10
         Top             =   120
         Width           =   3060
      End
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PENDAPATAN BERSIH"
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
      Left            =   8160
      TabIndex        =   44
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TGL. GAJI"
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
      Left            =   8160
      TabIndex        =   38
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label6 
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
      Left            =   8160
      TabIndex        =   23
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   8160
      TabIndex        =   19
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label Label5 
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
      TabIndex        =   17
      Top             =   2880
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
      TabIndex        =   16
      Top             =   1920
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
      TabIndex        =   15
      Top             =   1440
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
      TabIndex        =   13
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "prosespenggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cekstatus()
'cek status penggajihan
Set rsrekapgaji = New ADODB.Recordset
strsql = "select * from tabel_penggajihan where nik = '" & txtnik.Text & "' and bulan = '" & txtbulan.Text & "' and status = 'Selesai'"
rsrekapgaji.Open strsql, con, adOpenDynamic, adLockOptimistic
    If Not rsrekapgaji.EOF Then
        prosespenggajian.cmdsimpan.Visible = False
    Else
        prosespenggajian.cmdsimpan.Visible = True
    End If
End Sub

Private Sub simpan()
Set rsrekapgaji = New ADODB.Recordset
strsql = " update tabel_penggajihan set absen = '" & txtabsen.Text & "' " & _
        ",gapok = '" & txtgapok.Text & "' " & _
        ",tunj = '" & txttunjangan.Text & "' " & _
        ",potkes = '" & txtkesehatan.Text & "' " & _
        ",potket = '" & txtkerja.Text & "' " & _
        ",pajak = '" & txtpph.Text & "' " & _
        ",potkehadiran = '" & txtkehadiran.Text & "' " & _
        ",gatot = '" & txttotalpendapatan.Text & "' " & _
        ",totalpotongan = '" & txttotalpotongan.Text & "' " & _
        ",gajibersih = '" & txttotalbersih.Text & "' where nik = '" & txtnik.Text & "' and bulan = '" & txtbulan.Text & "'"
        Set rsrekapgaji = con.Execute(strsql, , adCmdText)
        MsgBox "Data Telah Diperbarui", vbOKOnly, "SUKSES"
        
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
transaksigajih.Refresh
bukakoneksi
bukaquerirekapgaji
gridquerirekapgaji
End Sub

Private Sub cmdbatal_Click()
Unload Me
MDIForm1.Enabled = True
MDIForm1.Show
End Sub

Private Sub cmdsimpan_Click()
kosong
End Sub

Private Sub Form_Load()
bukakoneksi
bukarekapgaji
cekstatus
End Sub

Private Sub txtbulan_Change()
Set rsrekapgaji = New ADODB.Recordset
strsql = "select * from tabel_penggajihan where nik = '" & txtnik.Text & "' and bulan = '" & txtbulan.Text & "'"
rsrekapgaji.Open strsql, con, adOpenDynamic, adLockOptimistic
If Not rsrekapgaji.EOF Then
txtgapok.Text = rsrekapgaji!gapok
txttunjangan.Text = rsrekapgaji!tunj
txtkesehatan.Text = rsrekapgaji!potkes
txtkerja.Text = rsrekapgaji!potket
txtpph.Text = rsrekapgaji!pajak
txtkehadiran.Text = rsrekapgaji!potkehadiran
txtabsen.Text = rsrekapgaji!absen
cekstatus
Else
End If
End Sub

'penjumlahan pendapatan
Private Sub txtgapok_Change()
txttotalpendapatan.Text = Val(Format(txtgapok.Text, "")) + Val(Format(txttunjangan.Text, ""))
txttotalpendapatan.Text = Format(Val(txttotalpendapatan.Text), "###,##")
End Sub
Private Sub txttunjangan_Change()
txttotalpendapatan.Text = Val(Format(txtgapok.Text, "")) + Val(Format(txttunjangan.Text, ""))
txttotalpendapatan.Text = Format(Val(txttotalpendapatan.Text), "###,##")
End Sub
'akhir penjumlahan pendapatan

'penjumlahan potongan
Private Sub txtkesehatan_Change()
txttotalpotongan.Text = Val(Format(txtkesehatan.Text, "")) + Val(Format(txtkerja.Text, "")) + Val(Format(txtpph.Text, "")) + Val(Format(txtkehadiran.Text, ""))
txttotalpotongan.Text = Format(Val(txttotalpotongan.Text), "###,##")
End Sub
Private Sub txtkerja_Change()
txttotalpotongan.Text = Val(Format(txtkesehatan.Text, "")) + Val(Format(txtkerja.Text, "")) + Val(Format(txtpph.Text, "")) + Val(Format(txtkehadiran.Text, ""))
txttotalpotongan.Text = Format(Val(txttotalpotongan.Text), "###,##")
End Sub
Private Sub txtpph_Change()
txttotalpotongan.Text = Val(Format(txtkesehatan.Text, "")) + Val(Format(txtkerja.Text, "")) + Val(Format(txtpph.Text, "")) + Val(Format(txtkehadiran.Text, ""))
txttotalpotongan.Text = Format(Val(txttotalpotongan.Text), "###,##")
End Sub
Private Sub txtkehadiran_Change()
txttotalpotongan.Text = Val(Format(txtkesehatan.Text, "")) + Val(Format(txtkerja.Text, "")) + Val(Format(txtpph.Text, "")) + Val(Format(txtkehadiran.Text, ""))
txttotalpotongan.Text = Format(Val(txttotalpotongan.Text), "###,##")
End Sub
'akhir penjumlahan potongan

'penjumlahan total bersih
Private Sub txttotalpendapatan_Change()
txttotalbersih.Text = Val(Format(txttotalpendapatan.Text, "")) - Val(Format(txttotalpotongan.Text, ""))
txttotalbersih.Text = Format(Val(txttotalbersih.Text), "###,##")
End Sub
Private Sub txttotalpotongan_Change()
txttotalbersih.Text = Val(Format(txttotalpendapatan.Text, "")) - Val(Format(txttotalpotongan.Text, ""))
txttotalbersih.Text = Format(Val(txttotalbersih.Text), "###,##")
End Sub
'akhir penjumlahan total bersih

Private Sub txtnik_Change()
Set rskaryawan = New ADODB.Recordset
strsql = "select * from tabel_karyawan where nik = '" & txtnik.Text & "'"
rskaryawan.Open strsql, con, adOpenDynamic, adLockOptimistic
If Not rskaryawan.EOF Then
txtalamat.Text = rskaryawan!alamat
Else
End If
End Sub

'format desimal
Private Sub txtgapok_LostFocus()
txtgapok.Text = Format(Val(txtgapok.Text), "###,##")
End Sub
Private Sub txttunjangan_LostFocus()
txttunjangan.Text = Format(Val(txttunjangan.Text), "###,##")
End Sub

Private Sub txtkesehatan_LostFocus()
txtkesehatan.Text = Format(Val(txtkesehatan.Text), "###,##")
End Sub
Private Sub txtkerja_LostFocus()
txtkerja.Text = Format(Val(txtkerja.Text), "###,##")
End Sub
Private Sub txtpph_LostFocus()
txtpph.Text = Format(Val(txtpph.Text), "###,##")
End Sub
Private Sub txtkehadiran_LostFocus()
txtkehadiran.Text = Format(Val(txtkehadiran.Text), "###,##")
End Sub


'set focus
Private Sub txtgapok_GotFocus()
txtgapok.SelStart = 0
txtgapok.SelLength = Len(txtgapok)
End Sub

Private Sub txttunjangan_GotFocus()
txttunjangan.SelStart = 0
txttunjangan.SelLength = Len(txttunjangan)
End Sub

Private Sub txtkesehatan_GotFocus()
txtkesehatan.SelStart = 0
txtkesehatan.SelLength = Len(txtkesehatan)
End Sub

Private Sub txtkerja_GotFocus()
txtkerja.SelStart = 0
txtkerja.SelLength = Len(txtkerja)
End Sub

Private Sub txtpph_GotFocus()
txtpph.SelStart = 0
txtpph.SelLength = Len(txtpph)
End Sub

Private Sub txtkehadiran_GotFocus()
txtkehadiran.SelStart = 0
txtkehadiran.SelLength = Len(txtkehadiran)
End Sub

Private Sub txtabsen_GotFocus()
txtabsen.SelStart = 0
txtabsen.SelLength = Len(txtabsen)
End Sub
'akhir set focus

Private Sub kosong()
If txtgapok.Text = "" Then
MsgBox "Gapok Tidak Boleh Kosong", vbOKOnly, "Peringatan"
txtgapok.SetFocus
Exit Sub
simpan
End If

If txtabsen.Text < 1 Then
MsgBox "Hari Kerja Tidak Boleh Kosong", vbOKOnly, "Peringatan"
txtabsen.SetFocus
Exit Sub
simpan
End If


If txttunjangan.Text = "" Then
txttunjangan.Text = Val(0)
End If

If txtkesehatan.Text = "" Then
txtkesehatan.Text = Val(0)
End If

If txtkerja.Text = "" Then
txtkerja.Text = Val(0)
End If

If txtpph.Text = "" Then
txtpph.Text = Val(0)
End If

If txtkehadiran.Text = "" Then
txtkehadiran.Text = Val(0)
End If

If txttotalpotongan.Text = "" Then
txttotalpotongan.Text = Val(0)
End If

simpan

End Sub
