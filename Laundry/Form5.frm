VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form5 
   Caption         =   "Member"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form5"
   ScaleHeight     =   5025
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4455
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtnama 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtalamat 
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtnotelp 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ID Pelanggan"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Telepon"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Pelanggan"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
   End
   Begin VB.TextBox txtcari 
      Height          =   405
      Left            =   9120
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton bsimpan 
      Caption         =   "SIMPAN"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton bclose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2655
      Left            =   4800
      TabIndex        =   0
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "                                                 MEMBER"
      BeginProperty Font 
         Name            =   "Garamond Premr Pro"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cari Data Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   15
      Top             =   960
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   4215
      Left            =   0
      Top             =   840
      Width           =   12735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
bsimpan.Enabled = False
bbatal.Enabled = False
cmdupdate.Enabled = False
End Sub
Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
End Sub

Sub bersih()
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Text = ""
Next
End Sub

Sub tampilgrid()
Call koneksi
rsmember.Open "select* from member order by idpelanggan", KON
Set Grid.DataSource = rsmember
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 3000
Grid.ColWidth(3) = 1000
Grid.ColWidth(4) = 1000
End Sub

Sub simpanmember()
simpan = "insert into member values('" & txtid.Text & "','" & txtnama.Text & "','" & txtalamat.Text & "','" & txtnotelp.Text & "')"
KON.Execute simpan
End Sub
Sub sqlmember()
SQL1 = "select * from member where namapelanggan like '%" & txtcari.Text & "%' order by namapelanggan asc"
KON.Execute SQL1
End Sub
Sub tampilmember()
Call koneksi
rsmember.Open "select* from member where namapelanggan like '%" & txtcari.Text & "%'", KON
Set Grid.DataSource = rsmember
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 3000
Grid.ColWidth(3) = 1000
Grid.ColWidth(4) = 1000
End Sub

Private Sub bbatal_Click()
Call bersih
bsimpan.Enabled = False
bbatal.Enabled = False
binput.Enabled = True
bclose.Enabled = True
cmdupdate.Enabled = False
End Sub

Private Sub bclose_Click()
If bclose.Caption = "CLOSE" Then
Unload Me
MDIForm1.Show
Call nonaktif
bclose.Caption = "CLOSE"
End If
End Sub

Private Sub binput_Click()
Call aktif
Call KodeOtomatis
txtid.Enabled = False
txtnama.SetFocus
binput.Enabled = False
bclose.Enabled = False
cmdupdate.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
End Sub
Sub KodeOtomatis()
If rsmember.State = adStateOpen Then rsmember.Close
rsmember.Open ("select * from member Where idpelanggan In(Select Max(idpelanggan)From member)Order By idpelanggan Desc"), KON, adOpenKeyset
    Dim Urutan As String * 5
    Dim Hitung As Long
        If rsmember.EOF Then
            Urutan = "MBR" + "01"
            txtid.Text = Urutan
        Else
Hitung = Right(rsmember!IDpelanggan, 2) + 1
Urutan = "MBR" & Right("00" & Hitung, 2)
        End If
        txtid.Text = Urutan
    
End Sub
Private Sub bsimpan_Click()
If txtid = "" Or txtnama = "" Or txtalamat = "" Or txtalamat = "" Then
MsgBox "Data Belum Lengkap"
Else
Call simpanmember
Call tampilgrid
Call bersih
bsimpan.Enabled = False
bsimpan.Enabled = False
binput.Enabled = True
bclose.Enabled = True
End If
MsgBox "Data Telah Tersimpan", vbInformation, "SIMPAN"
End Sub
Sub updateplg()
    Update = "UPDATE member SET namapelanggan = '" & txtnama.Text & "', alamat = '" & txtalamat.Text & "', telepon = '" & txtnotelp.Text & "' WHERE idpelanggan = '" & txtid.Text & "'"
    KON.Execute Update
End Sub

Private Sub cmdupdate_Click()
Call updateplg
Call bersih
Call tampilgrid
binput.Enabled = True
bclose.Enabled = True
bsimpan.Enabled = False
bbatal.Enabled = False
MsgBox "Data Telah Terupdate", vbInformation, "Update"
End Sub

Private Sub Form_Activate()
Call bersih
Call nonaktif
Call tampilgrid
End Sub

Private Sub Form_Load()
Call koneksi
End Sub

Private Sub Grid_DblClick()
    txtid.Text = Grid.TextMatrix(Grid.Row, 1)
    txtnama.Text = Grid.TextMatrix(Grid.Row, 2)
    txtalamat.Text = Grid.TextMatrix(Grid.Row, 3)
    txtnotelp.Text = Grid.TextMatrix(Grid.Row, 4)
    
    cmdupdate.Enabled = True
    binput.Enabled = False
    bclose.Enabled = False
    bsimpan.Enabled = False
    bbatal.Enabled = True
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
a = Grid.Row
plg = Grid.TextMatrix(a, 1)
Call koneksi
rsmember.Open "select * from member ", KON
With rsmember
If KeyAscii = 8 Then
If Not (.BOF And .EOF) Then
h = MsgBox("bener mau dihapus ?", vbQuestion + vbYesNo, "--Tanya--")
If h = vbYes Then
hapus = "delete from member where idpelanggan='" & plg & "'"
KON.Execute (hapus)
End If
End If
End If
End With
Call tampilgrid
Grid.Refresh
End Sub

Private Sub txtcari_Change()
Call koneksi
Call tampilmember
Call sqlmember
End Sub


