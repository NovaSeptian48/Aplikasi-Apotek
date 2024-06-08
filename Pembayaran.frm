VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   LinkTopic       =   "Form4"
   ScaleHeight     =   6285
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   4635
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CASH"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DEBIT"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QRIS"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   6345
      Left            =   0
      Picture         =   "Pembayaran.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Hide
Form5.Show
End Sub

Private Sub Command2_Click()
Form4.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub Command4_Click()
Form4.Hide
Form2.Show
End Sub
