VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMasterPelanggan 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   1080
      TabIndex        =   4
      Top             =   4200
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
      Left            =   5520
      Top             =   360
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode"
      Height          =   615
      Left            =   1200
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   615
      Left            =   1200
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "FormMasterPelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Command1.Caption = "INPUT" Then
    Command1.Caption = "SIMPAN"
    Call NoOtomatis
    Text2.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
        MsgBox "Silahkan isi data terlebih dahulu"
        Else
        
        Call bukaDB
        Dim TambahData As String
        TambahData = "Insert into tbl_pelanggan values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
        Koneksi.Execute TambahData
        MsgBox "Input Data Berhasil"
        Form_Activate
        End If
    End If
End Sub
Sub NoOtomatis()
Call bukaDB
    RSPelanggan.Open ("select * From tbl_pelanggan where kodepelanggan in(select(kodepelanggan) from tbl_pelanggan) order by kodepelanggan desc"), Koneksi
    RSPelanggan.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RSPelanggan
    If .EOF Then
        Urutan = "PLG" + "001"
        Text1 = Urutan
    Else
        Hitung = Right(RSPelanggan!KodePelanggan, 3) + 1
        Urutan = "PLG" + Right("000" & Hitung, 3)
    End If
    Text1 = Urutan
    End With
End Sub
Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim HapusData As String
HapusData = "delete from tbl_pelanggan where kodepelanggan = '" & Text1 & "'"
Koneksi.Execute HapusData
MsgBox "Hapus Data Berhasil"
Form_Activate
End If
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim EditData As String
EditData = "Update tbl_pelanggan set namapelanggan = '" & Text2 & "', alamatpelanggan = '" & Text3 & "', telppelanggan = '" & Text4 & "' where kodepelanggan = '" & Text1 & "'"
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
Text3.PasswordChar = ""
Text4 = ""
Call bukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "Select * from tbl_pelanggan"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Command1.Caption = "INPUT"
Command2.Caption = "EDIT"
Command3.Caption = "HAPUS"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call bukaDB
    RSPelanggan.Open "Select * From tbl_pelanggan where kodepelanggan = '" & Text1 & "'", Koneksi
    If RSPelanggan.EOF Then
    MsgBox "Data tidak ada"
Else
    Text1 = RSPelanggan!KodePelanggan
    Text2 = RSPelanggan!namapelanggan
    Text3 = RSPelanggan!alamatpelanggan
    Text4 = RSPelanggan!telppelanggan
End If
End If
End Sub

