VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMenuUtama 
   Caption         =   "Form Menu Utama Aplikasi Kasir"
   ClientHeight    =   5775
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   Picture         =   "FormMenuUtama.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4680
      Top             =   3120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   11
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "KODE"
            TextSave        =   "KODE"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "NAMA"
            TextSave        =   "NAMA"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "LEVEL"
            TextSave        =   "LEVEL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "TANGGAL"
            TextSave        =   "TANGGAL"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "JAM"
            TextSave        =   "JAM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1940
            MinWidth        =   1940
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "TERIMA KASIH ATAS KUNJUNGAN ANDA"
            TextSave        =   "TERIMA KASIH ATAS KUNJUNGAN ANDA"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu MenuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu Menu1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuKeluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu MenuMaster 
      Caption         =   "Master"
      Begin VB.Menu MenuKasir 
         Caption         =   "Kasir"
      End
      Begin VB.Menu MenuPelanggan 
         Caption         =   "Pelanggan"
      End
      Begin VB.Menu Menu2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuBarang 
         Caption         =   "Barang"
      End
   End
   Begin VB.Menu MenuTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu MenuPenjualan 
         Caption         =   "Penjualan"
      End
   End
   Begin VB.Menu MenuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu MenuLapPenjualan 
         Caption         =   "LapPenjualan"
      End
   End
   Begin VB.Menu MenuUtility 
      Caption         =   "Utility"
      Begin VB.Menu MenuGantiPassword 
         Caption         =   "GantiPassword"
      End
      Begin VB.Menu MenuManualBook 
         Caption         =   "ManualBook"
      End
   End
End
Attribute VB_Name = "FormMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call Terkunci
End Sub
Sub Terkunci()
MenuMaster.Visible = False
MenuTransaksi.Visible = False
MenuLaporan.Visible = False
MenuUtility.Visible = False
MenuLogout.Visible = False
MenuLogin.Visible = True

StatusBar1.Panels(2) = ""
StatusBar1.Panels(4) = ""
StatusBar1.Panels(6) = ""

End Sub

Private Sub MenuBarang_Click()
FormMenuBarang.Show vbModal
End Sub

Private Sub MenuKasir_Click()
FormMasterKasir.Show vbModal
End Sub

Private Sub MenuKeluar_Click()
End
End Sub

Private Sub MenuLogin_Click()
FormLogin.Show vbModal
End Sub

Private Sub MenuLogout_Click()
Call Terkunci
End Sub

Private Sub MenuPelanggan_Click()
FormMasterPelanggan.Show vbModal
End Sub

Private Sub MenuPenjualan_Click()
FormTransJual.Show vbModal
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(10) = Time$
StatusBar1.Panels(8) = Format(Date$, "dd-MM-yyyy")
End Sub
