VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "User"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   LinkTopic       =   "Form6"
   ScaleHeight     =   5025
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   3615
      Begin VB.TextBox txtpass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtnama 
         Height          =   405
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtuser 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cmbhakakses 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama User"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode User"
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
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hak Akses"
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
         TabIndex        =   9
         Top             =   2520
         Width           =   1815
      End
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton bsimpan 
      Caption         =   "SIMPAN"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton bclose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Caption         =   "                           USER"
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
      TabIndex        =   13
      Top             =   0
      Width           =   10695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   4215
      Left            =   0
      Top             =   840
      Width           =   10815
   End
End
Attribute VB_Name = "Form6"
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
cmbhakakses.Enabled = False
End Sub
Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
cmbhakakses.Enabled = True
End Sub
Sub bersih()
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Text = ""
Next
cmbhakakses.Text = ""
End Sub
Sub hakakses()
cmbhakakses.AddItem "Admin"
cmbhakakses.AddItem "User"
End Sub

Sub simpanadmin()
simpan = "insert into user values('" & txtuser.Text & "','" & txtnama.Text & _
"','" & txtpass.Text & "','" & cmbhakakses.Text & "')"
KON.Execute simpan
End Sub

Private Sub bbatal_Click()
Call bersih
cmbhakakses.Enabled = False
bsimpan.Enabled = False
bbatal.Enabled = False
binput.Enabled = True
bclose.Enabled = True
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
txtuser.Enabled = False
txtnama.SetFocus
binput.Enabled = False
bclose.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
If binput.Caption = "UPDATE" Then
End If
End Sub

Sub KodeOtomatis()
If rsuser.State = adStateOpen Then rsuser.Close
rsuser.Open ("select * from user Where kodeuser In(Select Max(kodeuser)From user)Order By kodeuser Desc"), KON, adOpenKeyset
    Dim Urutan As String * 6
    Dim Hitung As Long
    With rsuser
        If .EOF Then
            Urutan = "USR" + "01"
            txtuser = Urutan
        Else
Hitung = Right(rsuser!kodeuser, 2) + 1
Urutan = "USR" & Right("00" & Hitung, 2)
        End If
        txtuser = Urutan
    End With
End Sub
Private Sub bsimpan_Click()
If txtuser = "" Or txtnama = "" Or txtpass = "" Or cmbhakakses = "" Then
MsgBox "Data Belum Lengkap"
Else
Call simpanadmin
Call bersih
bsimpan.Enabled = False
bbatal.Enabled = False
binput.Enabled = True
bclose.Enabled = True
End If
MsgBox "Data Telah Tersimpan", vbInformation, "SIMPAN"
End Sub

Private Sub Form_Activate()
Call bersih
Call nonaktif
Call hakakses
End Sub

Private Sub Form_Load()
Call koneksi
End Sub



