VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11340
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Text            =   "* Kosongkan jika tidak tahu"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cbObat 
      Height          =   315
      ItemData        =   "order.frx":0000
      Left            =   2760
      List            =   "order.frx":0002
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Text            =   "Otomatis terisi"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ORDER"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   6345
      Left            =   0
      Picture         =   "order.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11400
   End
End
Attribute VB_Name = "Form2"
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
rs.Open "select * from TBLstock where obat='" & cbObat & "' and harga_jual", con, adOpenDynamic, adLockOptimistic, adCmdText
With rs
Text3.Text = !harga_jual
End With
End Sub
Private Sub Command1_Click()
connection
sql = "TBLorder"
Set rs = New ADODB.Recordset
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!id_user = Text1.Text
!obat = cbObat.Text
!jumlah = Text5.Text
!harga = Text3.Text
!total = Text4.Text
.Update
End With

connection
sql = "TBLriwayat"
Set rs = New ADODB.Recordset
rs.Open sql, con, adOpenDynamic, adLockOptimistic
With rs
.AddNew
!id_user = Text1.Text
!obat = cbObat.Text
!jumlah = Text5.Text
!harga = Text3.Text
!total = Text4.Text
.Update
End With
MsgBox "Order Sukses!!"
Form6.Text2.Text = Text4.Text
Form2.Hide
Form4.Show
End Sub

Private Sub Command2_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Form_Load()
connection
Set rs = New ADODB.Recordset
rs.Open "Select * from TBLstock where obat", con, adOpenDynamic, adLockOptimistic
With rs
Do Until .EOF
cbObat.AddItem ![obat]
.MoveNext
Loop
End With
Text2.Enabled = False
Text4.Enabled = False
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = Int(13) Then
    Dim A, B, C As Long
    A = Val(Text3.Text)
    B = Val(Text5.Text)
    
    C = A * B
    Text4.Text = C
End If
End Sub
