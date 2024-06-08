VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   LinkTopic       =   "Form3"
   Picture         =   "profile.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "HISTORY"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ORDER"
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOG OUT"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "profile.frx":3C66E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.connection
Public rs As ADODB.Recordset
Public sql As String
Public Function connection()
Set con = New ADODB.connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\apotek.mdb;"
End Function
Private Sub Command1_Click()
Form3.Hide
Form7.Show
Form7.Text1.Text = ""
Form7.Text2.Text = ""
End Sub

Private Sub Command2_Click()
Form2.Text1.Text = Text1.Text
Form3.Hide
Form2.Show
End Sub

Private Sub Command3_Click()

Form3.Hide
Form14.Show
End Sub
