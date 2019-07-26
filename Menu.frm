VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CV. AMDHAN PRINTING"
   ClientHeight    =   8520
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15105
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   4080
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8145
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Admin"
            TextSave        =   "Admin"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   18423
            MinWidth        =   8819
            Text            =   "Copy Right By @Feby 2019"
            TextSave        =   "Copy Right By @Feby 2019"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Time"
            TextSave        =   "Time"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   15045
      TabIndex        =   0
      Top             =   0
      Width           =   15105
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         DrawMode        =   4  'Mask Not Pen
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   6720
         Picture         =   "Menu.frx":0000
         ScaleHeight     =   1065
         ScaleWidth      =   1050
         TabIndex        =   4
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jl. Ir. H. Juanda No. 172 Margahayu - Bekasi Timur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   7320
         TabIndex        =   2
         Top             =   720
         Width           =   6495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CV. AMDHAN PRINTING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   555
         Left            =   6615
         TabIndex        =   1
         Top             =   120
         Width           =   7875
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu master 
      Caption         =   "Master"
      Begin VB.Menu user 
         Caption         =   "User"
      End
      Begin VB.Menu karyawan 
         Caption         =   "Karyawan"
      End
      Begin VB.Menu jabatan 
         Caption         =   "Jabatan"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu penggajihan 
         Caption         =   "Penggajihan"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jabatan_Click()
'buka menu jabatan
Unload Me
masterjabatan.Show
End Sub

Private Sub karyawan_Click()
'buka menu karyawan
Unload Me
masterkaryawan.Show
End Sub

Private Sub penggajihan_Click()
'buka menu hitung penggajihan
Unload Me
transaksigajih.Show
End Sub

Private Sub user_Click()
'buka menu pengguna
Unload Me
masteruser.Show
End Sub

Private Sub MDIForm_Load()
'set tanggal hari ini di status bar
Timer1.Interval = 1000
Timer1.Enabled = True
MDIForm1.StatusBar1.Panels(4) = Format(Date, "dd/MM/yy")
End Sub

Private Sub keluar_Click()
'perinta keluar aplikasi
Unload Me
End Sub
