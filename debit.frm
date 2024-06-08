VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
   LinkTopic       =   "Form6"
   ScaleHeight     =   6300
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "DONE"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "debit.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form6"
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
Dim A As Integer
A = MsgBox("Tidak ingin melanjutkan pembayaran?", vbQuestion + vbYesNo, "Cancel!")
If A = vbYes Then
Form2.Show
Form5.Hide
End If
End Sub

Private Sub Command2_Click()
connection
Set rs = New ADODB.Recordset
rs.Open "select* from TBLkeluar", con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!obat = Form2.cbObat.Text
!jumlah = Form2.Text5.Text
!harga = Form2.Text3.Text
.Update
End With

Form2.cbPenyakit.Refresh
Form2.cbObat.Refresh
Form2.Text3.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
MsgBox "Pembayaran sukses!"
Form6.Hide
Form2.Show
End Sub


Private Sub Form_Load()
Text2.Enabled = False
End Sub
