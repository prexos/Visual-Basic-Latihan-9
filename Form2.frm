VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   30
      Top             =   8400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1920
      TabIndex        =   18
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identitas Customer"
      Height          =   2655
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor"
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9240
      TabIndex        =   9
      Text            =   "January"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   11160
      TabIndex        =   8
      Text            =   "1999"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   360
      Left            =   2280
      TabIndex        =   5
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   360
      Left            =   8040
      TabIndex        =   3
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   360
      Left            =   10320
      TabIndex        =   2
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Height          =   360
      Left            =   10320
      TabIndex        =   1
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text12 
      Height          =   360
      Left            =   10320
      TabIndex        =   0
      Top             =   7440
      Width           =   2175
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
      Left            =   5400
      TabIndex        =   29
      Top             =   0
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   13320
      Y1              =   600
      Y2              =   600
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
      Left            =   5280
      TabIndex        =   28
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Nomor Faktur"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Tanggal Penjualan"
      Height          =   255
      Left            =   9240
      TabIndex        =   26
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   13320
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label8 
      Caption         =   "Nomor Stok"
      Height          =   255
      Left            =   600
      TabIndex        =   25
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Nama Stok"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Harga Jual"
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Unit Jual"
      Height          =   255
      Left            =   8520
      TabIndex        =   22
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Nilai Jual"
      Height          =   255
      Left            =   10920
      TabIndex        =   21
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Besaran Potongan"
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Nilai Penjualan Bersih"
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   7440
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
 TH = 1970
 Do While TH < 2022
 Combo3.AddItem TH
 TH = TH + 1
 Loop
 
 BLN = 1
 Do While BLN <= 12
 Combo2.AddItem (MonthName(BLN))
 BLN = BLN + 1
 Loop
 
 TG = 1
 Do While TG < 32
 Combo1.AddItem TG
 TG = TG + 1
 Loop
 
End Sub

