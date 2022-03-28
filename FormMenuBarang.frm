VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMenuBarang 
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2280
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   960
      TabIndex        =   8
      Top             =   4560
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6120
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Satuan Barang"
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jumlah Barang"
      Height          =   615
      Left            =   960
      TabIndex        =   12
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Harga Barang"
      Height          =   615
      Left            =   960
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Barang"
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Barang"
      Height          =   615
      Left            =   960
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FormMenuBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim TambahData As String
TambahData = "Insert into tbl_barang values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
Koneksi.Execute TambahData
MsgBox "Input Data Berhasil"
Form_Activate
End If

End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim HapusData As String
HapusData = "delete from tbl_barang where kodebarang = '" & Text1 & "'"
Koneksi.Execute HapusData
MsgBox "Hapus Data Berhasil"
Form_Activate
End If
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim EditData As String
EditData = "Update tbl_barang set namabarang = '" & Text2 & "', hargabarang = '" & Text3 & "', jumlahbarang = '" & Text4 & "',satuanbarang = '" & Text5 & "' where kodebarang = '" & Text1 & "'"
Koneksi.Execute EditData
MsgBox "Edit Data Berhasil"
Form_Activate
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call KondisiAwal
End Sub

Sub KondisiAwal()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Call bukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "Select * from tbl_barang"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call bukaDB
    RSBarang.Open "Select * From tbl_barang where kodebarang = '" & Text1 & "'", Koneksi
    If RSBarang.EOF Then
    MsgBox "Data tidak ada"
Else
    Text1 = RSBarang!kodebarang
    Text2 = RSBarang!namabarang
    Text3 = RSBarang!hargabarang
    Text4 = RSBarang!jumlahbarang
    Text5 = RSBarang!satuanbarang
End If
End If
End Sub


