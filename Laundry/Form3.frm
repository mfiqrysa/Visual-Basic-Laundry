VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form3"
   ScaleHeight     =   8100
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   6720
      TabIndex        =   33
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98828289
      CurrentDate     =   43242
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98828289
      CurrentDate     =   43242
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Ekspress Rp. 14.000,- Kg (1 hari)"
      Height          =   315
      Left            =   1440
      TabIndex        =   31
      Top             =   3240
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Standar Rp. 9.000,-  Kg (2 hari)"
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtnotrans 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtkasir 
      Height          =   285
      Left            =   6720
      TabIndex        =   13
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtnama 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtnotelp 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtalamat 
      Height          =   855
      Left            =   6000
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtharga 
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtqty 
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtsubtotal 
      Height          =   375
      Left            =   8880
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtbayar 
      Height          =   375
      Left            =   7920
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox txtkembali 
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton binput 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bsimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton bbatal 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton btutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   5880
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1720
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label14 
      Caption         =   "Pelayanan"
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
      TabIndex        =   29
      Top             =   2640
      Width           =   1215
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
      TabIndex        =   28
      Top             =   960
      Width           =   1815
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
      Left            =   6240
      TabIndex        =   27
      Top             =   4320
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
      Left            =   6240
      TabIndex        =   26
      Top             =   3720
      Width           =   1335
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
      Left            =   8880
      TabIndex        =   25
      Top             =   1560
      Width           =   1335
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
      Left            =   6480
      TabIndex        =   24
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Cucian"
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
      Left            =   4440
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
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
      Left            =   4560
      TabIndex        =   22
      Top             =   2640
      Width           =   1095
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
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
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
      TabIndex        =   20
      Top             =   1560
      Width           =   1935
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
      TabIndex        =   19
      Top             =   360
      Width           =   1095
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
      TabIndex        =   18
      Top             =   960
      Width           =   2055
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
      TabIndex        =   17
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lbayar 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   1080
      TabIndex        =   16
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6600
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
