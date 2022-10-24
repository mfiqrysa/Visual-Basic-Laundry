VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form4"
   ScaleHeight     =   5925
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   9360
      Top             =   3120
   End
   Begin VB.TextBox txtcari 
      Height          =   405
      Left            =   7080
      TabIndex        =   27
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmbsimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   1080
      TabIndex        =   26
      Top             =   5160
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MHSgrid 
      Height          =   1455
      Left            =   240
      TabIndex        =   25
      Top             =   3480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2566
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txttelp 
      Height          =   375
      Left            =   2280
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtplg 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtharga 
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtalamat 
      Height          =   855
      Left            =   1440
      TabIndex        =   19
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton btutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   3240
      TabIndex        =   18
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   2160
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtjenis 
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtjumlah 
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtpcs 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtkasir 
      Height          =   285
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtnotrans 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   43242
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   93388801
      CurrentDate     =   43242
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data Jasa Satuan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   22
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Terima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pelanggan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No Telpon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pcs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Pakaian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Cucian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Tanggal Selesai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
cmbsimpan.Enabled = False
bbatal.Enabled = False
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
txtkasir = MDIForm1.StatusBar1.Panels(3).Text
End Sub
Sub simpanjasa()
simpan = "insert into jasasatuan values('" & txtnotrans.Text & "','" & txtplg.Text & "','" & txtalamat.Text & "','" & txttelp.Text & _
"','" & DTPicker1.Value & "','" & DTPicker2.Value & "','" & txtpcs.Text & "','" & txtjumlah.Text & "','" & txtjenis.Text & "')"
KON.Execute simpan
End Sub

Sub sqlsatuan()
SQL1 = "select * from jasasatuan where namapelanggan like '%" & txtcari.Text & "%' order by namapelanggan asc"
KON.Execute SQL1

End Sub
Sub tampilsatuan()
Call koneksi
rssatuan.Open "select* from jasasatuan where namapelanggan like '%" & txtcari.Text & "%'", KON
Set Grid.DataSource = rssatuan
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 3000
Grid.ColWidth(3) = 1000
Grid.ColWidth(4) = 1000
Grid.ColWidth(5) = 1900
Grid.ColWidth(6) = 1000
Grid.ColWidth(7) = 1900
Grid.ColWidth(8) = 1000
Grid.ColWidth(9) = 1000
End Sub
Private Sub bbatal_Click()
Call bersih
cmbsimpan.Enabled = False
binput.Enabled = True
btutup.Enabled = True
bbatal.Enabled = False
End Sub

Private Sub binput_Click()
Call aktif
Call nomor
txtplg.SetFocus
txtid.Enabled = False
binput.Enabled = False
btutup.Enabled = False
cmbsimpan.Enabled = True
bbatal.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Call tampilgrid
If binput.Caption = "UPDATE" Then
End If
End Sub

Private Sub cmbedit_Click()
cmbedit.Caption = "&Update"
End Sub

Private Sub cmbsimpan_Click()
Call simpanjasa
Call tampilgrid
binput.Enabled = False
btutup.Enabled = False
End Sub

Private Sub Form_Load()
Call koneksi
End Sub


Sub tampilgrid()
Call koneksi
rssatuan.Open "select* from jasasatuan order by namapelanggan", KON
Set MHSgrid.DataSource = rssatuan
MHSgrid.ColWidth(0) = 0
MHSgrid.ColWidth(1) = 1600
MHSgrid.ColWidth(2) = 3000
MHSgrid.ColWidth(3) = 1000
MHSgrid.ColWidth(4) = 1000
MHSgrid.ColWidth(5) = 1900
MHSgrid.ColWidth(6) = 1000
MHSgrid.ColWidth(7) = 1900
MHSgrid.ColWidth(8) = 1000
MHSgrid.ColWidth(9) = 1000
End Sub
Private Sub btutup_Click()
If btutup.Caption = "TUTUP" Then
Unload Me
MDIForm1.Show
Call nonaktif
End If
End Sub

Private Sub Form_Activate()
txtkasir.Enabled = False
Call nonaktif
txtkasir = MDIForm1.StatusBar1.Panels(3).Text
End Sub


Sub nomor()
Dim cari As String
Call koneksi
rssatuan.Open "SELECT * FROM jasasatuan ORDER BY NOTRANSAKSI DESC ", KON
With rskiloan
If .EOF Then
txtnotrans = Format(Date, "yymm") + "001"
ElseIf Left(rssatuan!notransaksi, 4) <> Format(Date, "yymm") Then
txtnotrans = Format(Date, "yymm") + "001"
Else
No = .Fields("notransaksi") + 1
txtnotrans = Format(Date, "yymm") + Right("000" + No, 3)
End If
End With
End Sub

Private Sub MHSgrid_KeyPress(KeyAscii As Integer)
a = MHSgrid.Row
kode = MHSgrid.TextMatrix(a, 1)
Call koneksi
rssatuan.Open "select * from jasasatuan ", KON
With rssatuan
If KeyAscii = 8 Then
If Not (.BOF And .EOF) Then
h = MsgBox("bener mau dihapus ?", vbQuestion + vbYesNo, "--Tanya--")
If h = vbYes Then
hapus = "delete from jasasatuan where idpelanggan='" & kode & "'"
KON.Execute (hapus)
End If
End If
End If
End With
Call tampilgrid
MHSgrid.Refresh
End Sub

Private Sub txtcari_Change()
Call koneksi
Call tampilsatuan
Call sqlsatuan
End Sub
