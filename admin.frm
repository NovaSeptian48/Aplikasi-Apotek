VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   LinkTopic       =   "Form9"
   ScaleHeight     =   6315
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "LogOut"
      Height          =   435
      Left            =   9840
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SUPPLIER SERVICE"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OBAT KELUAR"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OBAT MASUK"
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STOK OBAT"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "admin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Hide
Form10.Show
End Sub

Private Sub Command2_Click()
Form9.Hide
Form11.Show
End Sub

Private Sub Command3_Click()
Form9.Hide
Form12.Show
End Sub

Private Sub Command4_Click()
Form9.Hide
Form15.Show
End Sub

Private Sub Command5_Click()
Form9.Hide
Form7.Show
Form7.Text1.Text = ""
Form7.Text2.Text = ""
End Sub
