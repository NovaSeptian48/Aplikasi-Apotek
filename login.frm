VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   LinkTopic       =   "Form7"
   ScaleHeight     =   6300
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btLogin 
      Caption         =   "LOGIN"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton btSign 
      Caption         =   "SIGN IN"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   6960
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   6960
      TabIndex        =   0
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   6345
      Left            =   0
      Picture         =   "login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.connection
Public rs As ADODB.Recordset
Private Sub btLogin_click()
Set rs = New ADODB.Recordset
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Username atau Password tidak boleh kosong!", vbExclamation, "Login Error!"
Exit Sub
End If
If rs.State = 1 Then
rs.Close
End If
rs.Open "select * from TBLuser where nama='" & Text1 & "' and password='" & Text2 & "' and status", con, adOpenDynamic, adLockOptimistic, adCmdText
If rs("status") = 1 Then
Form9.Show
Form7.Hide
Else
rs("status") = 2
Form7.Hide
Form3.Show
End If
If rs.EOF = True Then
MsgBox "Username atau Password salah!", vbCritical, "Login gagal"
Else
With rs
Form3.Text1.Text = "USR-" & !id
Form3.Text2.Text = !nama
Form3.Text3.Text = !email
Form3.Text4.Text = !alamat
Form3.Text5.Text = !no
Form3.Text9.Text = !kelamin
Form3.Text8.Text = !agama
Form3.Text10.Text = !telegram
End With
MsgBox "Login Success!!", vbInformation, "Success"
End If
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\apotek.mdb;"
End Sub

Private Sub btSign_click()
Form1.Show
Form7.Hide
End Sub

