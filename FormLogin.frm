VERSION 5.00
Begin VB.Form FormLogin 
   BackColor       =   &H0080FF80&
   Caption         =   "Form Login Aplikasi Kasir"
   ClientHeight    =   3615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   Picture         =   "FormLogin.frx":0000
   ScaleHeight     =   63.765
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   130.704
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "NewsGoth BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Kasir"
      BeginProperty Font 
         Name            =   "NewsGoth BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Terbuka()
FormMenuUtama.MenuMaster.Visible = True
FormMenuUtama.MenuTransaksi.Visible = True
FormMenuUtama.MenuLaporan.Visible = True
FormMenuUtama.MenuUtility.Visible = True
FormMenuUtama.MenuLogout.Visible = True
FormMenuUtama.MenuLogin.Visible = False
End Sub
Private Sub Command1_Click()
Call bukaDB
RSKasir.Open "Select * from tbl_kasir where kodekasir='" & Text1 & "' and passwordkasir= '" & Text2 & "'", Koneksi
If RSKasir.EOF Then
    MsgBox "kode kasir dan Password salah"
    Else
    Call Terbuka
    Me.Hide
    FormMenuUtama.StatusBar1.Panels(2) = RSKasir!KodeKasir
    FormMenuUtama.StatusBar1.Panels(4) = RSKasir!NamaKasir
    FormMenuUtama.StatusBar1.Panels(6) = RSKasir!LevelKasir
    
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Sub KondisiAwal()
Text1.MaxLength = 6
Text2.MaxLength = 30
Text2.PasswordChar = "x"
End Sub
Private Sub Form_Activate()
Text1 = "KSR001"
Text2 = "ADMIN"
End Sub

Private Sub Form_Load()
Call KondisiAwal
End Sub
