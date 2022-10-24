VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   Caption         =   "Pelayanan"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
      Begin VB.TextBox txtxnama 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtkode 
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtharga 
         Height          =   405
         Left            =   2760
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton bsimpan 
      Caption         =   "SIMPAN"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton bclose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtcari 
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2295
      Left            =   5040
      TabIndex        =   0
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "                                     PELAYANAN"
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
      TabIndex        =   14
      Top             =   0
      Width           =   11295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cari Data Pelayanan"
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
      Left            =   5520
      TabIndex        =   13
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   4815
      Left            =   0
      Top             =   960
      Width           =   11295
   End
End
Attribute VB_Name = "Form2"
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
rspelayanan.Open "select* from pelayanan order by kodepelayanan", KON
Set Grid.DataSource = rspelayanan
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 3000
Grid.ColWidth(3) = 1000
End Sub

Sub simpanpel()
simpan = "insert into pelayanan values('" & txtkode.Text & "','" & txtxnama.Text & "','" & txtharga.Text & "')"
KON.Execute simpan
End Sub
Sub sqlpel()
SQL1 = "select * from pelayanan where namapelayanan like '%" & txtcari.Text & "%' order by namapelayanan asc"
KON.Execute SQL1
End Sub
Sub tampilpel()
Call koneksi
rspelayanan.Open "select* from pelayanan where namapelayanan like '%" & txtcari.Text & "%'", KON
Set Grid.DataSource = rspelayanan
Grid.ColWidth(0) = 0
Grid.ColWidth(1) = 1600
Grid.ColWidth(2) = 3000
Grid.ColWidth(3) = 1000
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
txtkode.Enabled = False
txtxnama.SetFocus
binput.Enabled = False
bclose.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
cmdupdate.Enabled = False
End Sub
Sub KodeOtomatis()
If rspelayanan.State = adStateOpen Then rspelayanan.Close
rspelayanan.Open ("select * from pelayanan Where kodepelayanan In(Select Max(kodepelayanan)From pelayanan)Order By kodepelayanan Desc"), KON, adOpenKeyset
    Dim Urutan As String * 5
    Dim Hitung As Long
        If rspelayanan.EOF Then
            Urutan = "PYN" + "01"
            txtkode.Text = Urutan
        Else
Hitung = Right(rspelayanan!kodepelayanan, 2) + 1
Urutan = "PYN" & Right("00" & Hitung, 2)
        End If
        txtkode.Text = Urutan
    
End Sub

Private Sub bsimpan_Click()
If txtxnama = "" Or txtharga = "" Then
MsgBox "Data Belum Lengkap "
Else
Call simpanpel
Call tampilgrid
Call bersih
bsimpan.Enabled = False
bbatal.Enabled = False
binput.Enabled = True
bclose.Enabled = True
End If
MsgBox "Data Telah Tersimpan", vbInformation, "SIMPAN"
End Sub

Private Sub cmdupdate_Click()
Call updatepyn
Call bersih
Call tampilgrid
binput.Enabled = True
bclose.Enabled = True
bbatal.Enabled = False
bsimpan.Enabled = False
MsgBox "Data Telah Terupdate", vbInformation, "Update"
End Sub

Private Sub Form_Activate()
Call bersih
Call nonaktif
Call tampilgrid
End Sub
Sub updatepyn()
    Update = "UPDATE pelayanan SET namapelayanan = '" & txtxnama.Text & "', harga = '" & txtharga.Text & "' WHERE kodepelayanan = '" & txtkode.Text & "'"
    KON.Execute Update
End Sub
Private Sub Form_Load()
Call koneksi
End Sub

Private Sub Grid_DblClick()
 txtkode.Text = Grid.TextMatrix(Grid.Row, 1)
    txtxnama.Text = Grid.TextMatrix(Grid.Row, 2)
    txtharga.Text = Grid.TextMatrix(Grid.Row, 3)
    
    cmdupdate.Enabled = True
    binput.Enabled = False
    bclose.Enabled = False
    bsimpan.Enabled = False
    bbatal.Enabled = True
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
a = Grid.Row
pyn = Grid.TextMatrix(a, 1)
Call koneksi
rspelayanan.Open "select * from pelayanan ", KON
With rspelayanan
If KeyAscii = 8 Then
If Not (.BOF And .EOF) Then
h = MsgBox("bener mau dihapus ?", vbQuestion + vbYesNo, "--Tanya--")
If h = vbYes Then
hapus = "delete from pelayanan where kodepelayanan='" & pyn & "'"
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
Call tampilpel
Call sqlpel
End Sub

