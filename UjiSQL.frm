VERSION 5.00
Begin VB.Form UjiSQL 
   Caption         =   "Uji Coba SQL"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   10680
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox DataGrid1 
      Height          =   4440
      Left            =   120
      ScaleHeight     =   4380
      ScaleWidth      =   13260
      TabIndex        =   4
      Top             =   120
      Width           =   13320
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   5640
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4680
      Width           =   13320
   End
   Begin VB.PictureBox DT 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label TANGGAL 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Jumlah Data"
      Height          =   195
      Left            =   7080
      TabIndex        =   2
      Top             =   5760
      Width           =   885
   End
End
Attribute VB_Name = "UjiSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
TANGGAL = Date
On Error Resume Next
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBDealer.mdb"
DT.RecordSource = "select * from bayarcicilan where idkredit='xxx'"
Set DataGrid1.DataSource = DT
DataGrid1.Refresh
Text1 = ""
Text1.SetFocus
Command1.Default = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
    Dim Pesan As String
    On Error GoTo salah
    DT.RecordSource = Text1
    DT.Refresh
    Text2 = DT.Recordset.RecordCount
    
    If DT.Recordset.EOF Then
        Pesan = MsgBox("Data Tidak Ditemukan")
        DT.Refresh
        Text1.SetFocus
    End If
    On Error GoTo 0
    Exit Sub
salah:
    MsgBox "Syntax SQL Salah..!"
End Sub


