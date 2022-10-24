VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form7 
   BackColor       =   &H00808000&
   Caption         =   "Laporan"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5520
   LinkTopic       =   "Form7"
   ScaleHeight     =   6525
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " HARIAN"
      Height          =   1095
      Left            =   600
      TabIndex        =   11
      Top             =   840
      Width           =   3735
      Begin VB.ComboBox cmbhari 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Hari"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MINGGUAN"
      Height          =   1935
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   3735
      Begin VB.ComboBox cmbawal 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cmbakhir 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Minggu Awal"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Minggu Akhir"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "BULANAN"
      Height          =   1815
      Left            =   600
      TabIndex        =   1
      Top             =   4320
      Width           =   3735
      Begin VB.ComboBox cmbbulan 
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cmbtahun 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Bulan"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tahun"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   4560
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   4560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbakhir_Click()
If cmbawal = " " Then
MsgBox "tanggal awal kosong", , "informasi"
cmbawal.SetFocus
Exit Sub
End If
CrystalReport2.SelectionFormula = "{transaksi.tgltransaksi} in date (" & cmbawal.Text & _
") to date (" & cmbakhir.Text & ")"
CrystalReport2.ReportFileName = App.Path & "\laporanmingguan.rpt"
CrystalReport2.WindowState = crptMaximized
CrystalReport2.RetrieveDataFiles
CrystalReport2.Action = 1
End Sub

Private Sub cmbawal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmbbulan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmbhari_Click()
CrystalReport2.SelectionFormula = "totext ({transaksi.tgltransaksi})='" & cmbhari & "'"
CrystalReport2.ReportFileName = App.Path & "\laporanharian.rpt"
CrystalReport2.WindowState = crptMaximized
CrystalReport2.RetrieveDataFiles
CrystalReport2.Action = 1
End Sub

Private Sub cmbhari_KeyPress(KeyAscii As Integer)
If cmbhari = " " Or KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmbtahun_Click()
Call koneksi
rstransaksi.Open "select * from transaksi where month(tgltransaksi)=' " & Val(cmbbulan) & _
" ' and year(tgltransaksi)=' " & (cmbtahun) & " ' ", KON
If rstransaksi.EOF Then
MsgBox "data tidak ditemukan"
Exit Sub
cmbbulan.SetFocus
End If
CrystalReport2.SelectionFormula = "Month({transaksi.tgltransaksi})=" & Val(cmbbulan.Text) & _
" and Year ({transaksi.tgltransaksi})=" & Val(cmbtahun.Text)
CrystalReport2.ReportFileName = App.Path & "\laporanbulanan.rpt"
CrystalReport2.WindowState = crptMaximized
CrystalReport2.RetrieveDataFiles
CrystalReport2.Action = 1

End Sub

Private Sub cmdclose_Click()
Unload Me
MDIForm1.Show
End Sub



Private Sub Form_Load()
Call koneksi
rstransaksi.Open "select distinct tgltransaksi from transaksi order by 1", KON
rstransaksi.Requery
Do Until rstransaksi.EOF
cmbhari.AddItem rstransaksi!tgltransaksi
cmbawal.AddItem Format(rstransaksi!tgltransaksi, "YYYY ,MM, DD")
cmbakhir.AddItem Format(rstransaksi!tgltransaksi, "YYYY ,MM, DD")
rstransaksi.MoveNext
Loop
For i = 1 To 12
cmbbulan.AddItem i
Next i
For i = 10 To 20
cmbtahun.AddItem 2000 + i
Next i
End Sub

