VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Customer 
   Caption         =   "Data Customer Kelompok 6"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7080
      OleObjectBlob   =   "Customer.frx":0000
      Top             =   2400
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":0234
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   350
      Left            =   1800
      TabIndex        =   14
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   2400
      Width           =   5535
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   4320
         TabIndex        =   3
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
      Begin VB.CommandButton Cmdedit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdinput 
         Caption         =   "&Input"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.TextBox Text5 
      Height          =   350
      Left            =   1800
      TabIndex        =   12
      Top             =   1680
      Width           =   1500
   End
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   4000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   4000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3413
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3600
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin VB.CheckBox C3 
      Caption         =   "Slip Gaji"
      Height          =   350
      Left            =   3360
      TabIndex        =   10
      Top             =   2040
      Width           =   1500
   End
   Begin VB.CheckBox C2 
      Caption         =   "KK"
      Height          =   350
      Left            =   3360
      TabIndex        =   9
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CheckBox C1 
      Caption         =   "KTP"
      Height          =   350
      Left            =   3360
      TabIndex        =   8
      Top             =   1320
      Width           =   1500
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":02AC
      TabIndex        =   16
      Top             =   600
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":0312
      TabIndex        =   17
      Top             =   960
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":037C
      TabIndex        =   18
      Top             =   1320
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":03E8
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   375
      Left            =   240
      OleObjectBlob   =   "Customer.frx":044A
      TabIndex        =   20
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub AutoNumber()
Call BukaDB
RSCustomer.Open ("SELECT * FROM CUSTOMER WHERE KODECUS in(select max(KODECUS) from CUSTOMER)order by KODECUS desc"), CONN
RSCustomer.Requery
     Dim Urutan As String * 6
    Dim hitung As Long
    With RSCustomer
        If .EOF Then
            Urutan = "KC" + "01"
            Text1 = Urutan
        Else
            hitung = Right(!Kodecus, 3) + 1
            Urutan = "KC" + Right("0" & hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub


Private Sub Cmdtutup_Click()
    Select Case Cmdtutup.Caption
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\DBDealer.mdb"
Adodc1.RecordSource = "customer"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd

Call AutoNumber
End Sub

Sub Form_Load()
Call BukaDB
Text1.MaxLength = 6
Text2.MaxLength = 50
Text3.MaxLength = 50
Text4.MaxLength = 15
Text5.MaxLength = 15
Text6.MaxLength = 15
KondisiAwal
End Sub

Private Sub KosongkanText()
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    C1.Value = vbUnchecked
    C2.Value = vbUnchecked
    C3.Value = vbUnchecked
    Text6 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = False
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    C1.Enabled = True
    C2.Enabled = True
    C3.Enabled = True
    Text6.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
    C1.Enabled = False
    C2.Enabled = False
    C3.Enabled = False
    Text6.Enabled = False
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
    With RSCustomer
        If Not RSCustomer.EOF Then
            Text2 = RSCustomer!Nama
            Text3 = RSCustomer!Alamat
            Text4 = RSCustomer!Telepon
            Text5 = RSCustomer!HP
            If RSCustomer!ktp <> 0 Then
                C1.Value = vbChecked
            Else
                C1.Value = vbUnchecked
            End If
            
            If RSCustomer!KK <> 0 Then
                C2.Value = vbChecked
            Else
                C2.Value = vbUnchecked
            End If
            
            If RSCustomer!slipgaji <> 0 Then
                C3.Value = vbChecked
            Else
                C3.Value = vbUnchecked
            End If
            
            Text6 = RSCustomer!KETERANGAN
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
        
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Customer (KodeCus,Nama,Alamat,Telepon,HP,KTP,KK,Slipgaji,Keterangan) values " & _
            "('" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & C1.Value & "','" & C2.Value & "','" & C3.Value & "','" & Text6 & "')"
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
        Text1.SetFocus
    Else
        If Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Customer Set Nama= '" & Text2 & "', " & _
            "Alamat='" & Text3 & "', " & _
            "Telepon='" & Text4 & "', " & _
            "HP='" & Text5 & "', " & _
            "KTP='" & C1 & "', " & _
            "KK='" & C2 & "', " & _
            "slipgaji='" & C3 & "', " & _
            "keterangan = '" & Text6 & "' where KodeCus='" & Text1 & "'"
            CONN.Execute SQLEdit
            Form_Activate
            Call KondisiAwal
        End If
    End If
End Sub

Private Sub Cmdhapus_Click()
    If Cmdhapus.Caption = "&Hapus" Then
        Cmdinput.Enabled = False
        Cmdedit.Enabled = False
        Cmdtutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text1.SetFocus
    End If
End Sub


Function CariData()
    Call BukaDB
    RSCustomer.Open "Select * From Customer where KodeCus='" & Text1 & "'", CONN
End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    If Len(Text1) < 4 Then
        MsgBox "Kode Harus 4 Digit"
        Text1.SetFocus
        Exit Sub
    Else
        Text2.SetFocus
    End If

    If Cmdinput.Caption = "&Simpan" Then
        Call CariData
        If Not RSCustomer.EOF Then
            TampilkanData
            MsgBox "Kode Customer Sudah Ada"
            KosongkanText
            Text1.SetFocus
        Else
            Text2.SetFocus
        End If
    End If
    
    If Cmdedit.Caption = "&Simpan" Then
        Call CariData
        If Not RSCustomer.EOF Then
            TampilkanData
            Text1.Enabled = False
            Text2.SetFocus
        Else
            MsgBox "Kode Customer Tidak Ada"
            Text1 = ""
            Text1.SetFocus
        End If
    End If
    
    If Cmdhapus.Enabled = True Then
        Call CariData
        If Not RSCustomer.EOF Then
            TampilkanData
            Pesan = MsgBox("Yakin akan dihapus", vbYesNo + vbCritical)
            If Pesan = vbYes Then
                Dim SQLHapus As String
                SQLHapus = "Delete From Customer where kodeCus= '" & Text1 & "'"
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
    If KeyAscii = 13 Then Text5.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub Text5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then Text6.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub


Private Sub Text6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cmdinput.Enabled = True Then
            Cmdinput.SetFocus
        ElseIf Cmdedit.Enabled = True Then
            Cmdedit.SetFocus
        End If
    End If

End Sub


