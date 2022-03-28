VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormTransJual 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   13050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6720
      Top             =   1080
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5880
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label18 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kembali"
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dibayar"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label LBLNamaKasir 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   10200
      TabIndex        =   14
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label LBLJam 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   10200
      TabIndex        =   13
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label LBLTanggal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   10200
      TabIndex        =   12
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kasir"
      Height          =   495
      Left            =   8520
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jam"
      Height          =   495
      Left            =   8520
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label LBLTelp 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3360
      Width           =   6255
   End
   Begin VB.Label LBLAlamat 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   6255
   End
   Begin VB.Label LBLNama 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label LBLNoJual 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pelanggan"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Jual"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "FormTransJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Call bukaDB
RSPelanggan.Open "select * from tbl_Pelanggan where kodepelanggan='" & Combo1 & "'", Koneksi
    If Not RSPelanggan.EOF Then
        LBLNama = RSPelanggan!NamaPelanggan
        LBLAlamat = RSPelanggan!AlamatPelanggan
        LBLTelp = RSPelanggan!TelpPelanggan
    End If
    Koneksi.Close
End Sub

Private Sub Form_Load()
Call NoOtomatis
LBLNamaKasir = FormMenuUtama.StatusBar1.Panels(4)

Call bukaDB
RSPelanggan.Open "select * from tbl_Pelanggan", Koneksi
    Combo1.Clear
    Do Until RSPelanggan.EOF
        Combo1.AddItem RSPelanggan!KodePelanggan
        RSPelanggan.MoveNext
    Loop
    
End Sub


Sub NoOtomatis()
Call bukaDB
    RSJual.Open ("select * From tbl_jual where NoJual in(select(NoJual) from tbl_jual) order by NoJual desc"), Koneksi
    RSJual.Requery
    Dim Urutan As String * 12
    Dim Hitung As Long
    With RSJual
    If .EOF Then
        Urutan = "J" + "00000000001"
        LBLNoJual = Urutan
    Else
        Hitung = Right(RSJual!NoJual, 11) + 1
        Urutan = "J" + Right("00000000000" & Hitung, 11)
    End If
    LBLNoJual = Urutan
    End With
    
End Sub

Private Sub Timer1_Timer()
LBLJam = Time$
LBLTanggal = Format(Date$, "dd-MM-yyyy")
End Sub
