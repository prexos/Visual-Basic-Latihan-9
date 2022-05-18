VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      TabIndex        =   31
      Top             =   8280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   3480
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Customer"
      Height          =   2655
      Left            =   1800
      TabIndex        =   12
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   9840
      TabIndex        =   11
      Text            =   "18"
      Top             =   2040
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   10800
      TabIndex        =   10
      Text            =   "Mei"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   360
      Left            =   12720
      TabIndex        =   9
      Text            =   "2022"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   2640
      Width           =   5055
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   6480
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   360
      Left            =   9600
      TabIndex        =   4
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   360
      Left            =   11880
      TabIndex        =   3
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Height          =   360
      Left            =   11880
      TabIndex        =   2
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      Height          =   360
      Left            =   11880
      TabIndex        =   1
      Top             =   7680
      Width           =   2175
   End
   Begin VB.ComboBox Combo4 
      Height          =   360
      Left            =   8640
      TabIndex        =   0
      Text            =   "Rabu"
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "CV BITFINEX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   30
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1560
      X2              =   14880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "FAKTUR PENJUALAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Faktur"
      Height          =   255
      Left            =   2040
      TabIndex        =   28
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Tanggal Penjualan"
      Height          =   255
      Left            =   10800
      TabIndex        =   27
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1560
      X2              =   14880
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label8 
      Caption         =   "Nomor Stok"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Nama Stok"
      Height          =   255
      Left            =   4560
      TabIndex        =   25
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Harga Jual"
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Unit Jual"
      Height          =   255
      Left            =   10080
      TabIndex        =   23
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Nilai Jual"
      Height          =   255
      Left            =   12480
      TabIndex        =   22
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Besaran Potongan"
      Height          =   255
      Left            =   10080
      TabIndex        =   21
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Nilai Penjualan Bersih"
      Height          =   255
      Left            =   9840
      TabIndex        =   20
      Top             =   7680
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
 Text5.Text = Combo4.Text + Combo1.Text + " " + Combo2.Text + " " + Combo3.Text
End Sub

Private Sub Combo2_Click()
 Text5.Text = Combo4.Text + Combo1.Text + " " + Combo2.Text + " " + Combo3.Text
End Sub

Private Sub Combo3_Click()
 Text5.Text = Combo4.Text + Combo1.Text + " " + Combo2.Text + " " + Combo3.Text
End Sub

Private Sub Combo4_Click()
 Text5.Text = Combo4.Text + Combo1.Text + " " + Combo2.Text + " " + Combo3.Text
End Sub

Private Sub Command1_Click()
 End
End Sub

Private Sub Form_Activate()
 'interface
 Form2.WindowState = 2
 Text1.SetFocus
 
 'Hari
 nhari = " Senin Senin Selasa Rabu   Kamis  Jumat  Sabtu  Minggu "
 hr = 1
 Do While hr < 8
 hrf = Mid(nhari, 7 * hr, 7)
 Combo4.AddItem hrf
 hr = hr + 1
 Loop
 
 'Tanggal
 tgl = 1
 Do While tgl < 32
 Combo1.AddItem tgl
 tgl = tgl + 1
 Loop
 'Bulan
 bln = 1
 Do While bln < 13
 Combo2.AddItem MonthName(bln)
 bln = bln + 1
 Loop
 'Tahun
 thn = 1920
 Do While thn < 2023
 Combo3.AddItem thn
 thn = thn + 1
 Loop
 
End Sub
