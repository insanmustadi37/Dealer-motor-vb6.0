VERSION 5.00
Begin VB.Form Pemakai 
   Caption         =   "Data Pemakai Aplikasi"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form3"
   ScaleHeight     =   4260
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ADO 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   400
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   400
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   850
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "X"
      TabIndex        =   4
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   3540
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   2000
   End
   Begin VB.TextBox KodeDasar 
      Height          =   350
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2000
   End
   Begin VB.PictureBox DG 
      Height          =   1845
      Left            =   120
      ScaleHeight     =   1785
      ScaleWidth      =   4995
      TabIndex        =   11
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pemakai"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "Pemakai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Koneksi
ADO.ConnectionString = PathData
ADO.RecordSource = "pemakai"
ADO.Refresh
Set DG.DataSource = ADO
DG.Refresh
End Sub

Sub Form_Load()
    Call Koneksi
    Text1.MaxLength = 5
    Text2.MaxLength = 30
    Text3.MaxLength = 15
    KondisiAwal
    Combo1.AddItem "USER"
    Combo1.AddItem "ADMINISTRATOR"
End Sub

Private Sub combo1_click()
Text1.Enabled = False
 If Combo1 = "USER" Then
        KodeDasar = "USR"
        Call KODEOTO
    ElseIf Combo1 = "ADMINISTRATOR" Then
        KodeDasar = "ADM"
        Call KODEOTO
    End If
End Sub

Private Sub combo1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    If Combo1 = "" Then
        MsgBox "Status harus diisi"
        Combo1.SetFocus
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub dg_Keypress(Keyascii As Integer)
If Keyascii = 13 Then
        If CmdSimpan.Enabled = False Then
        MsgBox "pilih dulu command Edit atau Hapus"
        Exit Sub
    End If

    If CmdEdit.Enabled = True Then
        Combo1.Enabled = False
        Text1.Enabled = False
        Combo1 = DG.Columns(3)
        Text1 = DG.Columns(0)
        Text2 = DG.Columns(1)
        Text3 = DG.Columns(2)
        'Text2.SetFocus
    End If
    
    If CmdHapus.Enabled = True Then
        Combo1 = DG.Columns(3)
        Text1 = DG.Columns(0)
        Text2 = DG.Columns(1)
        Text3 = DG.Columns(2)
        Call CariData
        If Not RSPemakai.EOF Then
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                hapus = "delete * from pemakai where kodepmk='" & Text1 & "'"
                Conn.Execute hapus
                Call KondisiAwal
            Else
                Call KondisiAwal
            End If
        End If
    End If
End If
End Sub

Private Sub KODEOTO()
Call Koneksi
RSPemakai.Open "SELECT count(statuspmk) as ketemu FROM pemakai where statuspmk='" & Combo1 & "'", Conn
RSPemakai.Requery
If RSPemakai!ketemu = 0 Then
    Text1 = KodeDasar + "1"
Else
    Hitung = RSPemakai!ketemu + 1
    Text1 = KodeDasar + Right("0" & Hitung, 1)
End If
End Sub

Function CariData()
    Call Koneksi
    RSPemakai.Open "Select * From Pemakai where KodePmk='" & Text1 & "'", Conn
End Function

Private Sub CmdBatal_Click()
KosongkanText
TidakSiapIsi
KondisiAwal
End Sub

Private Sub CmdSimpan_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
    MsgBox "Data Belum Lengkap...!"
    Exit Sub
ElseIf Len(Text3) < 6 Then
    MsgBox "Password minimal 6 karakter"
    Text3.SetFocus
    Exit Sub
End If
    
If CmdInput.Enabled = True Then
    Dim SQLTambah1 As String
    SQLTambah1 = "Insert Into Pemakai (KodePmk,NamaPmk,PassPmk,StatusPmk) values " & _
    "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Combo1 & "')"
    Conn.Execute SQLTambah1
Else
    Dim SQLEdit As String
    SQLEdit = "Update Pemakai Set NamaPmk= '" & Text2 & "', PassPmk = '" & Text3 & "',StatusPmk = '" & Combo1 & "' where KodePmk='" & Text1 & "'"
    Conn.Execute SQLEdit
End If
Form_Activate
KondisiAwal

End Sub

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Combo1 = ""
    KodeDasar = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Combo1.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Combo1.Enabled = False
End Sub

Private Sub KondisiAwal()
KosongkanText
TidakSiapIsi
CmdInput.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
CmdSimpan.Enabled = False
CmdBatal.Enabled = False
CmdTutup.Enabled = True
Form_Activate
End Sub

Private Sub TampilkanData()
On Error Resume Next
Text2 = RSPemakai!NamaPmk
Text3 = RSPemakai!PassPMK
Combo1 = RSPemakai!StatusPmk
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdSimpan.Enabled = True
        CmdBatal.Enabled = True
        CmdTutup.Enabled = False
        SiapIsi
        KosongkanText
        Combo1.SetFocus
    End If
End Sub

Private Sub CmdEdit_Click()
If CmdEdit.Caption = "&Edit" Then
    CmdInput.Enabled = False
    CmdHapus.Enabled = False
    CmdTutup.Enabled = False
    CmdSimpan.Enabled = True
    CmdBatal.Enabled = True
    SiapIsi
    Combo1.Enabled = False
    Text1.SetFocus
End If
End Sub

Private Sub CmdHapus_Click()
If CmdHapus.Caption = "&Hapus" Then
    CmdTutup.Enabled = False
    CmdInput.Enabled = False
    CmdEdit.Enabled = False
    CmdSimpan.Enabled = True
    CmdBatal.Enabled = True
    SiapIsi
    Text1.SetFocus
End If

End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    Combo1.Enabled = False
    If Left(Text1, 3) <> "ADM" And Left(Text1, 3) <> "USR" Then
        MsgBox "Tiga Digit Pertama Harus ADM atau USR"
        Text1.SetFocus
        Exit Sub
    End If

    If CmdInput.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                MsgBox "Kode Pemakai Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode Pemakai Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSPemakai.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Pemakai where KodePmk= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Form_Activate
                    KondisiAwal
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub text2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub text3_keypress(Keyascii As Integer)
Text3.PasswordChar = "X"
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdSimpan.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdSimpan.SetFocus
        End If
    End If

End Sub

