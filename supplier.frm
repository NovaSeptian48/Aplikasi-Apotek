VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   LinkTopic       =   "Form15"
   ScaleHeight     =   6315
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "SUPPLY"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
   End
   Begin VB.ComboBox cbObat 
      Height          =   315
      Left            =   7680
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox cbSup 
      Height          =   315
      ItemData        =   "supplier.frx":0000
      Left            =   2520
      List            =   "supplier.frx":000D
      TabIndex        =   5
      Text            =   "Select ID"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox txtHarga 
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtJumlah 
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtNadmin 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtNsup 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "supplier.frx":0029
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form15"
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

Private Sub cbObat_Click()
connection
Set rs = New ADODB.Recordset
rs.Open "select * from TBLsup where obat='" & cbObat & "' and harga", con, adOpenDynamic, adLockOptimistic, adCmdText
With rs
txtHarga.Text = !harga
End With
End Sub

Private Sub cbSup_Click()
Select Case cbSup.Text
Case "SUP-01"
txtNsup.Text = "Ridwan Utomo"
Case "SUP-02"
txtNsup.Text = "Aryeswara"
Case "SUP-03"
txtNsup.Text = "M. Zafir"
End Select
End Sub


Private Sub Command1_Click()
Form15.Hide
Form9.Show
End Sub

Private Sub Command2_Click()
connection
sql = "TBLsupplier"
Set rs = New ADODB.Recordset
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!id_supplier = cbSup.Text
!nama_supplier = txtNsup.Text
!nama_admin = txtNadmin.Text
!nama_obat = cbObat.Text
!jumlah = txtJumlah.Text
!harga = txtHarga.Text
.Update
End With
MsgBox "Permintaan Supply Berhasil!! Menunggu pengiriman", vbInformation, "Success!"
End Sub

Private Sub Form_Load()
connection
Set rs = New ADODB.Recordset
rs.Open "Select * from TBLsup where obat", con, adOpenDynamic, adLockOptimistic
With rs
Do Until .EOF
cbObat.AddItem ![obat]
.MoveNext
Loop
End With
End Sub

