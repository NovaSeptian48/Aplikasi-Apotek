VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtLahir 
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   137756673
      CurrentDate     =   45280
   End
   Begin VB.ComboBox cbAgama 
      Height          =   315
      ItemData        =   "registerr.frx":0000
      Left            =   7560
      List            =   "registerr.frx":0016
      TabIndex        =   9
      Text            =   "Pilih"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.ComboBox cbKelamin 
      Height          =   315
      ItemData        =   "registerr.frx":004C
      Left            =   7560
      List            =   "registerr.frx":0059
      TabIndex        =   8
      Text            =   "Pilih"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton btSgn 
      Caption         =   "SIGN IN"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton btLgn 
      Caption         =   "LOG IN"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox txtTele 
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtNomor 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox txtAlamat 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "registerr.frx":007F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   5880
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
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

Private Sub btLgn_Click()
Form7.Show
Form1.Hide
End Sub

Private Sub btSgn_click()
If txtUser.Text = "" Or txtEmail.Text = "" Or txtPass.Text = "" Or txtAlamat.Text = "" Or txtNomor.Text = "" Or cbKelamin.Text = "" Or cbAgama.Text = "" Or cbStatus.Text = "" Or txtTele.Text = "" Then
MsgBox "Tidak boleh ada field yang kosong!", vbExclamation, "Field kosong"
Else
connection
sql = "TBLuser"
Set rs = New ADODB.Recordset
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!nama = txtUser.Text
!email = txtEmail.Text
!Password = txtPass.Text
!alamat = txtAlamat.Text
!no = txtNomor.Text
!kelamin = cbKelamin.Text
!lahir = dtLahir.Value
!agama = cbAgama.Text
!Status = 2
!telegram = txtTele.Text
.Update
End With
MsgBox "Registered!!", vbInformation, "Success!"
Form1.Hide
Form7.Show
End If
End Sub


