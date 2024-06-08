VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   LinkTopic       =   "Form11"
   Picture         =   "obatmasuk.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHjual 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox txtHbeli 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3720
      Width           =   4335
   End
   Begin VB.TextBox txtJumlah 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtJenis 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox txtNama 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1680
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TAMBAHKAN"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "obatmasuk.frx":2B561
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form11"
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
Form11.Hide
Form9.Show
End Sub

Private Sub Command2_Click()
connection
sql = "TBLstock"
Set rs = New ADODB.Recordset
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!obat = txtNama.Text
!harga_jual = txtHjual.Text
!harga_beli = txtHbeli.Text
!jumlah = txtJumlah.Text
!jenis = txtJenis.Text
.Update
End With
MsgBox "Obat ditambahkan ke stock", vbInformation, "Success!"
Form1.Hide
End Sub
