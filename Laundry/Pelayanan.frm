VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Login"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tuser 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox tpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmbcancel 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmblogin 
      BackColor       =   &H00404000&
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      FillColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   -120
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "                      LOGIN"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   9600
      Left            =   -480
      Top             =   -1680
      Width           =   9600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbcancel_Click()
keluar = MsgBox("Anda Yakin Ingin Keluar?", vbQuestion + vbYesNo, "Keluar?")
    If keluar = vbYes Then
    Unload Me
    End If
End Sub

Private Sub cmblogin_Click()
Call koneksi
rsuser.Open "select * from user where username ='" & tuser & _
"'and password='" & tpass & "'", KON
If rsuser.EOF Then
MsgBox "Password anda salah", vbCritical
tpass.Text = ""
tpass.SetFocus
Else
Me.Visible = False
Unload Me
MDIForm1.Show
MDIForm1.StatusBar1.Panels(4) = rsuser!hakakses
If MDIForm1.StatusBar1.Panels(4) = "Admin" Then
MDIForm1.mnutama.Enabled = True
MDIForm1.transaksi.Enabled = True
MDIForm1.laporan.Enabled = True
ElseIf MDIForm1.StatusBar1.Panels(4) = "User" Then
MDIForm1.mnutama.Enabled = False
MDIForm1.transaksi.Enabled = True
MDIForm1.laporan.Enabled = True
End If
End If
End Sub



Private Sub Form_Activate()
tuser.Enabled = True
tpass.Enabled = True
cmblogin.Enabled = True
tuser.SetFocus
tuser.MaxLength = 5
tpass.MaxLength = 8
End Sub


Private Sub tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmblogin.Enabled = True
cmblogin.SetFocus
End If
End Sub

Private Sub tuser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi
rsuser.Open "select * from user where username='" & tuser.Text & "'", KON
If rsuser.EOF Then
MsgBox "username tidak ditemukan,silakan masukan username lainnya", vbCritical
tuser.Text = Clear
Else
tpass.Enabled = True
tpass.SetFocus
MDIForm1.StatusBar1.Panels(3) = rsuser!UserName
cmblogin.Enabled = True
tuser.Enabled = False
End If
End If
End Sub
