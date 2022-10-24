VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   Caption         =   "Transaksi"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form3"
   ScaleHeight     =   8745
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      TabIndex        =   49
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Express Rp.14.000,-Kg (1 Hari)"
      Height          =   375
      Left            =   5280
      TabIndex        =   48
      Top             =   3480
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Standar Rp.9.000,-Kg (2 Hari)"
      Height          =   435
      Left            =   5280
      TabIndex        =   47
      Top             =   2880
      Width           =   2655
   End
   Begin VB.TextBox txtkdpyn 
      Height          =   405
      Left            =   6840
      TabIndex        =   46
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtjenisply 
      Height          =   405
      Left            =   6840
      TabIndex        =   43
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtharga 
      Height          =   405
      Left            =   6840
      TabIndex        =   41
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtdiskon 
      Height          =   375
      Left            =   10320
      TabIndex        =   38
      Top             =   2880
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   8160
      TabIndex        =   28
      Top             =   3480
      Width           =   3975
      Begin VB.TextBox txtkembali 
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtbayar 
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lbayar 
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1680
         TabIndex        =   34
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Kembali"
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
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Bayar"
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
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdlayanan 
      Caption         =   "List Pelayanan"
      Height          =   495
      Left            =   4560
      TabIndex        =   27
      Top             =   5640
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   26
      Top             =   4920
      Width           =   4335
   End
   Begin VB.TextBox txtkasir 
      Height          =   405
      Left            =   10320
      TabIndex        =   25
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txttgl 
      Height          =   375
      Left            =   10320
      TabIndex        =   24
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox cmbid 
      Height          =   315
      Left            =   2280
      TabIndex        =   23
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtnotelp 
      Height          =   405
      Left            =   2280
      TabIndex        =   12
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtnama 
      Height          =   405
      Left            =   2280
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtnotrans 
      Height          =   405
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtberat 
      Height          =   405
      Left            =   6720
      TabIndex        =   7
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtjumlah 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtjenis 
      Height          =   405
      Left            =   6720
      TabIndex        =   5
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton bsimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   8160
      Width           =   1095
   End
   Begin VB.CommandButton btutup 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   8160
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   8160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   2778
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98697217
      CurrentDate     =   43242
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98697217
      CurrentDate     =   43242
   End
   Begin VB.Label Label21 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Pelayanan"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   45
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Transaksi"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8880
      TabIndex        =   44
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Pelayanan"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   42
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   40
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label17 
      Caption         =   "%"
      Height          =   255
      Left            =   11160
      TabIndex        =   39
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Diskon"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8880
      TabIndex        =   37
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000D&
      Caption         =   "                                           TRANSAKSI"
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
      Height          =   975
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   12255
   End
   Begin VB.Label Label6 
      Caption         =   "KG"
      Height          =   255
      Left            =   7560
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No Transaksi"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Terima"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   8880
      TabIndex        =   19
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "No Telpon"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Cucian"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Cucian"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Selesai"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub simpanTEMP()
simpan = "insert into TEMP() values('" & cmbid.Text & "','" & txtnama.Text & "','" & txtnotelp.Text & "','" & DTPicker1 & _
"','" & DTPicker2 & "','" & txtberat.Text & "','" & txtjenis.Text & "','" & txtjumlah.Text & "','" & Val(lbayar) & "')"
KON.Execute (simpan)
End Sub

Private Sub btutup_Click()
Call hapusTEMP
Unload Me
MDIForm1.Show
End Sub

Sub cetak()
CrystalReport1.SelectionFormula = "{transaksi.notransaksi}='" & txtnotrans & "'"
CrystalReport1.ReportFileName = App.Path & "\struk1.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1
End Sub
Private Sub cmbid_Click()
Call koneksi
Set rsmember = New ADODB.Recordset
rsmember.Open "select*from member where idpelanggan='" & Left(cmbid.Text, 5) & "'", _
KON, adOpenDynamic, adLockOptimistic
rsmember.Requery
With rsmember
If .EOF And .BOF Then
MsgBox "idpelanggan Tidak ada", _
vbOKOnly + vbCritical, "Error"
Exit Sub
Else
cmbid.Text = rsmember!IDpelanggan
txtnama.Text = rsmember!namapelanggan
txtnotelp.Text = rsmember!telepon
txtberat.SetFocus
txtdiskon.Enabled = True
End If
End With
rsmember.Close
End Sub
Sub nope()
rsmember.Open " select * from member", KON
cmbid.Clear
Do While Not rsmember.EOF
cmbid.AddItem rsmember!IDpelanggan
rsmember.MoveNext
Loop
End Sub
Private Sub cmdlayanan_Click()
List1.Visible = True
End Sub

Private Sub Form_Activate()
txtnotrans.Enabled = False
txtdiskon.Enabled = False
Call isilist
Call semula
Call nope
List1.Visible = False
txtkasir = MDIForm1.StatusBar1.Panels(3).Text
End Sub
Sub nomor()
txttgl = Format(Date, "DD/MM/YYYY")
Dim cari As String
Call koneksi
rstransaksi.Open "SELECT * FROM transaksi ORDER BY NOTRANSAKSI DESC ", KON
With rstransaksi
If .EOF Then
txtnotrans = Format(Date, "yymm") + "001"
ElseIf Left(rstransaksi!notransaksi, 4) <> Format(Date, "yymm") Then
txtnotrans = Format(Date, "yymm") + "001"
Else
no = .Fields("notransaksi") + 1
txtnotrans = Format(Date, "yymm") + Right("000" + no, 3)
End If
End With
End Sub
Private Sub Form_Load()
Call koneksi
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
a = grid.Row
kode = grid.TextMatrix(a, 1)
Call koneksi
rstemp.Open "select * from TEMP ", KON
With rstemp
If KeyAscii = 8 Then
If Not (.BOF And .EOF) Then
h = MsgBox("bener mau dihapus ?", vbQuestion + vbYesNo, "--Tanya--")
If h = vbYes Then
hapus = "delete from TEMP where notransaksi='" & kode & "'"
Set rstemp = KON.Execute(hapus)
txtnama.Text = ""
txtnotelp.Text = ""
txtberat.Text = ""
txtjumlah.Text = ""
txtjenis.Text = ""
End If
End If
End If
End With
Call tampilgrid
grid.Refresh

End Sub

Sub semula()
Call bersih
Call nonaktif
bsimpan.Enabled = False
binput.Enabled = True
btutup.Enabled = True
bbatal.Enabled = False
Option1.Value = False
Option2.Value = False
End Sub

Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
cmdlayanan.Enabled = False
cmbid.Enabled = False
txtkasir.Enabled = False
bsimpan.Enabled = False
bbatal.Enabled = False
End Sub

Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
cmbid.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
End Sub

Sub bersih()
Dim kontrol As Control
lbayar.Caption = "Bayar"
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Text = ""
Next
cmbid.Text = ""
txtkasir = MDIForm1.StatusBar1.Panels(3).Text
End Sub
Sub simpantransaksi()
txttgl = Format(Date, "YYYY/MM/DD")
simpan = "insert into transaksi() values('" & txtnotrans.Text & _
"','" & DTPicker1 & "','" & DTPicker2.Value & "','" & txtberat.Text & "','" & txtjenis.Text & "','" & Text1.Text & "','" & txtdiskon.Text & "','" & txtbayar.Text & "','" & txtkembali.Text & "','" & txttgl.Text & "','" & txtkasir.Text & "')"
KON.Execute (simpan)
End Sub
Private Sub bbatal_Click()
Call semula
Option1.Value = False
Option2.Value = False
End Sub

Private Sub binput_Click()
Call aktif
cmdlayanan.Enabled = True
txtdiskon.Enabled = False
cmbid.Enabled = True
txtnotrans.Enabled = False
btutup.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
binput.Enabled = False
Call hapusTEMP
Call bikinTEMP
Call tampilgrid
Call nomor
txtnama.SetFocus
End Sub
Private Sub bsimpan_Click()
Call simpantransaksi
Call simpandetailjual
x = MsgBox("cetak?", vbYesNo, "cetak")
If x = vbYes Then
Call cetak
Call tampilgrid
Call semula
Else
Call tampilgrid
Call semula
End If
Call koneksi
End Sub

Sub tampilgrid()
Call koneksi
rstemp.Open "select * from TEMP", KON
Set grid.DataSource = rstemp
grid.ColWidth(0) = 100
grid.ColWidth(1) = 1500
grid.ColWidth(2) = 1700
grid.ColWidth(3) = 1700
grid.ColWidth(4) = 1700
grid.ColWidth(5) = 1500
grid.ColWidth(6) = 1500
grid.ColWidth(7) = 1500
grid.ColWidth(8) = 1500
grid.ColWidth(9) = 1500
grid.TextMatrix(0, 1) = "idpelanggan"
grid.TextMatrix(0, 2) = "namapelanggan"
grid.TextMatrix(0, 3) = "notelpon"
grid.TextMatrix(0, 4) = "tanggalterima"
grid.TextMatrix(0, 5) = "tanggalselesai"
grid.TextMatrix(0, 6) = "beratcucian"
grid.TextMatrix(0, 7) = "jeniscucian"
grid.TextMatrix(0, 8) = "qty"
grid.TextMatrix(0, 9) = "subtotal"
End Sub
Sub bikinTEMP()
bikin = "create table TEMP(idpelanggan varchar(6),namapelanggan varchar(30),notelpon varchar(12),tanggalterima varchar(11),tanggalselesai varchar(11),beratcucian int,jeniscucian varchar (30),qty int, subttl double)"
KON.Execute (bikin)
Call tampilgrid
End Sub
Sub hapusTEMP()
hapus = "drop table if exists TEMP"
KON.Execute (hapus)
End Sub
Sub simpandetailjual()
Dim simpan, fak, nmplg As String
Dim jumlah As Integer
Dim subtotal As Double

For a = 1 To (grid.Rows - 1)
fak = txtnotrans
IDpelanggan = grid.TextMatrix(a, 1)
quantity = grid.TextMatrix(a, 8)
subtotal = grid.TextMatrix(a, 9)

simpan = "insert into detailtransaksi()values('" & fak & "','" & IDpelanggan & "','" & txtkdpyn & _
"','" & quantity & "','" & Val(subtotal) & "')"
Set rsdetail = KON.Execute(simpan)
Next a
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
lbayar = Val(lbayar) + Val(9000)
Text1.Text = "Standar"
txtbayar.SetFocus
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
lbayar = Val(lbayar) + Val(14000)
Text1.Text = "Express"
txtbayar.SetFocus
End If
End Sub

Private Sub List1_Click()
hrg = "select * from pelayanan where kodepelayanan='" & Left(List1, 7) & "'"
Set rspelayanan = KON.Execute(hrg)
txtkdpyn.Text = rspelayanan!kodepelayanan
txtjenisply.Text = rspelayanan!namapelayanan
txtharga.Text = rspelayanan!harga
txtjenis.SetFocus
List1.Visible = False
End Sub
Sub isilist()
rspelayanan.Open "select * from pelayanan", KON
List1.Clear
Do While Not rspelayanan.EOF
List1.AddItem rspelayanan!kodepelayanan & Space(10) & rspelayanan!namapelayanan & Space(3) & rspelayanan!harga
rspelayanan.MoveNext
Loop
End Sub

Private Sub txtbayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(lbayar) > Val(txtbayar) Then
MsgBox "uang bayar kurang"
txtbayar.SetFocus
txtkembali.Enabled = False
Else
txtkembali.Enabled = True
txtkembali = Val(txtbayar) - Val(lbayar)
bsimpan.SetFocus
End If
End If
End Sub

Private Sub txtberat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi
rspelayanan.Open "select * from pelayanan where kodepelayanan='" & txtkdpyn & "'", KON
If Val(txtberat) > 100 Then
MsgBox "stok kurang"
txtberat.SetFocus
Exit Sub
Else
lbayar = Val(txtharga) * Val(txtberat)
Call simpanTEMP
Call tampilgrid
Call isilist
ttl = 0
For a = 1 To (grid.Rows - 1)
x = Val(grid.TextMatrix(a, 9))
ttl = ttl + x
Next a
lbayar.Caption = ttl
t = MsgBox("Mau Tambah Pelayanan Lagi?", vbQuestion + vbYesNo, "konfirmasi")
If t = vbYes Then
txtkdpyn.Text = ""
txtjenisply.Text = ""
txtharga.Text = ""
txtberat.Text = ""
Else
ambilstok = False
Me.Refresh
grid.Refresh
txtbayar.SetFocus
End If
End If
End If
End Sub
Private Sub txtdiskon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
lbayar = Val(lbayar) * Val(100 - txtdiskon) / 100
End If
End Sub

Private Sub txtidpel_Change()

End Sub

Private Sub txtjumlah_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi
rspelayanan.Open "select * from pelayanan where kodepelayanan='" & txtkdpyn & "'", KON
If Val(txtberat) > 100 Then
MsgBox "stok kurang"
txtberat.SetFocus
Exit Sub
Else
lbayar = Val(txtharga) * Val(txtjumlah)
Call simpanTEMP
Call tampilgrid
Call isilist
ttl = 0
For a = 1 To (grid.Rows - 1)
x = Val(grid.TextMatrix(a, 9))
ttl = ttl + x
Next a
lbayar.Caption = ttl
t = MsgBox("Mau Tambah Pembelian Lagi?", vbQuestion + vbYesNo, "konfirmasi")
If t = vbYes Then
List1.Visible = True
txtkdpyn.Text = ""
txtjenisply.Text = ""
txtharga.Text = ""
txtberat.Text = ""
txtjumlah.Text = ""
txtjenis.Text = ""
txtberat.Text = ""
Else
ambilstok = False
Me.Refresh
grid.Refresh
txtbayar.SetFocus
End If
End If
End If
End Sub
