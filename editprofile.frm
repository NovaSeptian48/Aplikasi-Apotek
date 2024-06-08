VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
      Height          =   255
      Left            =   6240
      TabIndex        =   26
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton CommandUpdate 
      Caption         =   "DONE"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CheckBox CheckWA 
      Caption         =   "Centang jika nomor HP anda sudah terdaftar di whatsapp"
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox TextTelegram 
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox ComboAgama 
      Height          =   315
      Left            =   8280
      TabIndex        =   10
      Text            =   "Pilih"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox ComboTanggalLahir 
      Height          =   315
      Left            =   8280
      TabIndex        =   9
      Text            =   "Pilih"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox TextHomorHP 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   4080
      Width           =   2775
   End
   Begin VB.ComboBox ComboKelamin 
      Height          =   315
      Left            =   8280
      TabIndex        =   7
      Text            =   "Pilih"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox ComboStatus 
      Height          =   315
      Left            =   8280
      TabIndex        =   6
      Text            =   "Pilih"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox TextAlamat 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox TextPassword 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox TextEmail 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox TextNama 
      Height          =   405
      Left            =   2640
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox TextID 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor HP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Whatsapp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   23
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Telegram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Agama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Lahir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EDIT PROFILE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4913
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text5_Change()

End Sub

Private Sub Command1_Click()

End Sub
