VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormMasterKasir 
   Caption         =   "Form Master Kasir"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   5640
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "INPUT"
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   480
      TabIndex        =   8
      Top             =   4320
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
      Left            =   4920
      Top             =   480
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1920
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level"
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "FormMasterKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim TambahData As String
TambahData = "Insert into tbl_kasir values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "')"
Koneksi.Execute TambahData
MsgBox "Input Data Berhasil"
Form_Activate
End If

End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim HapusData As String
HapusData = "delete from tbl_kasir where kodekasir = '" & Text1 & "'"
Koneksi.Execute HapusData
MsgBox "Hapus Data Berhasil"
Form_Activate
End If
End Sub

Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Silahkan isi data terlebih dahulu"
Else

Call bukaDB
Dim EditData As String
EditData = "Update tbl_kasir set namakasir = '" & Text2 & "', passwordkasir = '" & Text3 & "', levelkasir = '" & Combo1 & "' where kodekasir = '" & Text1 & "'"
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
Text3.PasswordChar = "x"
combo = ""

Combo1.Clear
Combo1.AddItem "ADMIN"
Combo1.AddItem "USER"
Call bukaDB
Adodc1.ConnectionString = Koneksi
Adodc1.RecordSource = "Select kodekasir,namakasir,levelkasir from tbl_kasir"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call bukaDB
    RSKasir.Open "Select * From tbl_kasir where kodekasir = '" & Text1 & "'", Koneksi
    If RSKasir.EOF Then
    MsgBox "Data tidak ada"
Else
    Text1 = RSKasir!KodeKasir
    Text2 = RSKasir!NamaKasir
    Text3 = RSKasir!PasswordKasir
    Combo1 = RSKasir!LevelKasir
End If
End If
End Sub
