VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Motor 
   Caption         =   "Data Motor Kelompok 6"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6360
      OleObjectBlob   =   "Motor.frx":0000
      Top             =   960
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6255
      Begin VB.TextBox Text4 
         Height          =   350
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1250
      End
      Begin VB.TextBox Text3 
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1250
      End
      Begin VB.TextBox Text2 
         Height          =   350
         Left            =   1200
         TabIndex        =   5
         Top             =   600
         Width           =   4260
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   1250
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3480
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "Motor.frx":0234
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "Motor.frx":029A
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "Motor.frx":0300
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "Motor.frx":0368
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   6255
      Begin VB.CommandButton Cmdinput 
         Caption         =   "&Input"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1000
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2805
      Left            =   0
      TabIndex        =   10
      Top             =   2760
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4948
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Motor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Auto()
Call BukaDB
RSMotor.Open ("SELECT * FROM MOTOR WHERE KODEMTR in(select max(KODEMTR) from MOTOR)order by KODEMTR desc"), CONN
RSMotor.Requery
     Dim Urutan As String * 6
    Dim hitung As Long
    With RSMotor
        If .EOF Then
            Urutan = "KM" + "001"
            Text1 = Urutan
        Else
            hitung = Right(!Kodemtr, 3) + 1
            Urutan = "KM" + Right("00" & hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub


Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBDealer.mdb"
Adodc1.RecordSource = "Motor"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd
Call Auto
End Sub

Sub Form_Load()
Call BukaDB
Text1.MaxLength = 5
Text2.MaxLength = 10
Text3.MaxLength = 10
Text4.MaxLength = 8
KondisiAwal
End Sub

Private Sub KosongkanText()
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    Cmdinput.Caption = "&Input"
    Cmdedit.Caption = "&Edit"
    Cmdhapus.Caption = "&Hapus"
    Cmdtutup.Caption = "&Batal"
    Cmdinput.Enabled = True
    Cmdedit.Enabled = True
    Cmdhapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSMotor
        If Not RSMotor.EOF Then
            Text2 = RSMotor!merk
            Text3 = RSMotor!warna
            Text4 = RSMotor!harga
        End If
    End With
End Sub
Private Sub Cmdinput_Click()
    If Cmdinput.Caption = "&Input" Then
        Cmdinput.Caption = "&Simpan"
        Cmdedit.Enabled = False
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Motor (KodeMtr,Merk,Warna,Harga) values ('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "')"
            CONN.Execute SQLTambah
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub


Private Sub Cmdedit_Click()
    If Cmdedit.Caption = "&Edit" Then
        Cmdinput.Enabled = False
        Cmdedit.Caption = "&Simpan"
        Cmdhapus.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
    Else
        If Text2 = "" Or Text3 = "" Or Text4 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Motor Set Merk= '" & Text2 & "', Warna='" & Text3 & "', Harga='" & Text4 & "'  where KodeMtr='" & Text1 & "'"
            CONN.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub Cmdhapus_Click()
 If Cmdhapus.Caption = "&Hapus" Then
        Cmdinput.Enabled = False
        Cmdhapus.Caption = "&Simpan"
        Cmdedit.Enabled = False
        Cmdtutup.Caption = "&Batal"
        SiapIsi
        Text1.SetFocus
    Else
        Pesan = MsgBox("Yakin akan dihapus", vbYesNo + vbCritical)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From Motor where kodeMtr= '" & Text1 & "'"
                CONN.Execute SQLHapus
                KondisiAwal
                Form_Activate
       End If
       End If
End Sub

Private Sub Cmdtutup_Click()
    Select Case Cmdtutup.Caption
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Function CariData()
    Call BukaDB
    RSMotor.Open "Select * From Motor where KodeMtr='" & Text1 & "'", CONN
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Len(Text1) < 5 Then
        MsgBox "Kode Harus 5 Digit"
        Text1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            MsgBox "Kode Motor Sudah Ada"
            KosongkanText
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            Text1.Enabled = False
            Text2.SetFocus
        Else
            MsgBox "Kode Motor Tidak Ada"
            Text1 = ""
            Text1.SetFocus
        End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
        If Not RSMotor.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From Motor where kodeMtr= '" & Text1 & "'"
                CONN.Execute SQLHapus
                KondisiAwal
                Form_Activate
            Else
                KondisiAwal
                Cmdhapus.SetFocus
            End If
        Else
            MsgBox "Data Tidak ditemukan"
            Text1.SetFocus
        End If
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Text4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0

End Sub


