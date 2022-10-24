VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Master"
   ClientHeight    =   10650
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16710
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   9240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   9855
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   9795
      ScaleWidth      =   16650
      TabIndex        =   1
      Top             =   0
      Width           =   16710
      Begin VB.Image Image1 
         Height          =   4305
         Left            =   0
         Picture         =   "MDIForm1.frx":8A25C
         Top             =   -11040
         Width           =   7875
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   10155
      Width           =   16710
      _ExtentX        =   29475
      _ExtentY        =   873
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2011
            MinWidth        =   2011
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2187
            MinWidth        =   2187
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   3000
   End
   Begin VB.Menu mnutama 
      Caption         =   "Master"
      Begin VB.Menu MNUser 
         Caption         =   "User"
      End
      Begin VB.Menu MNMember 
         Caption         =   "Member"
      End
      Begin VB.Menu MNDaftar 
         Caption         =   "Pelayanan"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "Transaksi"
   End
   Begin VB.Menu laporan 
      Caption         =   "Laporan"
      Begin VB.Menu MNLaporan 
         Caption         =   "Laporan"
      End
      Begin VB.Menu laporanpelayanan 
         Caption         =   "Laporan Transaksi"
      End
   End
   Begin VB.Menu logout 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub laporanpelayanan_Click()
CrystalReport1.ReportFileName = App.Path & "\laporantransaksi1.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1
End Sub

Private Sub logout_Click()
keluar = MsgBox("Anda Yakin Ingin Log Out?", vbQuestion + vbYesNo, "Keluar?")
    Unload Me
Form1.Show 'Perintah Menampilkan Form 1
MDIForm1.Visible = False 'Menyembunyikan mdiform1
End Sub

Private Sub MNDaftar_Click()
Form2.Show
End Sub
Private Sub MNLaporan_Click()
Form7.Show
End Sub
Private Sub MNMember_Click()
Form5.Show
End Sub
Private Sub MNUser_Click()
Form6.Show
End Sub
Private Sub Timer1_Timer()
StatusBar1.Panels(1) = Time()
StatusBar1.Panels(2) = Date
End Sub

Private Sub transaksi_Click()
Form3.Show
End Sub
