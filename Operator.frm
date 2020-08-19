VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Operator 
   Caption         =   "Data Pemakai Aplikasi Kelompok 6"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form3"
   ScaleHeight     =   5160
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   1680
      Width           =   6135
      Begin VB.CommandButton cmdtambah 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdtutup 
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   1000
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1000
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1920
      OleObjectBlob   =   "Operator.frx":0000
      Top             =   4440
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   3540
   End
   Begin VB.TextBox Text3 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "X"
      TabIndex        =   0
      Top             =   1080
      Width           =   2000
   End
   Begin MSAdodcLib.Adodc ADO 
      Height          =   375
      Left            =   120
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBDealer.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=DBDealer.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "OPERATOR"
      Caption         =   "ado"
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
   Begin MSDataGridLib.DataGrid DG 
      Bindings        =   "Operator.frx":0234
      Height          =   1845
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "KodeOpr"
         Caption         =   "Kode"
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
         DataField       =   "NamaOpr"
         Caption         =   "Nama"
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
      BeginProperty Column02 
         DataField       =   "passwordopr"
         Caption         =   "Password"
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
      BeginProperty Column03 
         DataField       =   "StatusOPr"
         Caption         =   "Status"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1995,024
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Operator.frx":0246
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Operator.frx":02B0
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Operator.frx":0326
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   375
      Left            =   120
      OleObjectBlob   =   "Operator.frx":039C
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub kosong()
Text2 = ""
Text3 = ""
Text4 = ""
End Sub
Private Sub isi()
Text2.Enabled = True
Text3.Enabled = True
End Sub
Private Sub blmisi()
Text2.Enabled = False
Text3.Enabled = False
End Sub


Private Sub awalform()
blmisi
cmdtambah.Enabled = False
Cmdhapus.Enabled = False
Cmdtutup.Enabled = False
End Sub


Private Sub Cmdhapus_Click()
If MsgBox("Apakah yakin ingin hapus?", vbQuestion + vbYesNo, "NOTICE") = vbYes Then
ADO.Recordset.Delete
DG.Refresh
End If
Call kosong
Call awalform
End Sub


Private Sub Cmdsimpan_Click()
ADO.Recordset!NAMAOPR = Text2
ADO.Recordset!PASSWORDOPR = Text3
ADO.Recordset!STATUSOPR = Combo1
ADO.Recordset.Update
DG.Refresh
kosong
awalform

End Sub

Private Sub cmdtambah_Click()
If Text2 = ADO.Recordset!NAMAOPR Then
MsgBox "Username telah digunakan!", vbCritical, "ERROR"
Text2.SetFocus
ElseIf Text2 = "" Then
MsgBox "Username masih kosong!", , ""
Text2.SetFocus
ElseIf Text2 = "" Then
MsgBox "Password masih kosong!", , ""
Text3.SetFocus
ElseIf Combo1 = "Pilih---" Then
MsgBox "Status belum dipilih!", , ""
Combo1.SetFocus
Else
ADO.Recordset.AddNew
ADO.Recordset!NAMAOPR = Text2
ADO.Recordset!PASSWORDOPR = Text3
ADO.Recordset!KODEOPR = Text1
ADO.Recordset!STATUSOPR = Combo1
ADO.Recordset.Update
DG.Refresh
kosong

End If
End Sub

Private Sub Cmdtutup_Click()
Call awalform
Call kosong

End Sub



Private Sub DG_Click()
isi
cmdtambah.Enabled = False
Cmdhapus.Enabled = True
Cmdtutup.Enabled = True
Text1.Enabled = False
Combo1.Enabled = False
If ADO.Recordset.EOF And ADO.Recordset.BOF Then
MsgBox "Data tidak ditemukan !"
Else
Text2.Text = ADO.Recordset!NAMAOPR
Text3.Text = ADO.Recordset!PASSWORDOPR
Combo1 = ADO.Recordset!STATUSOPR
Text1 = ADO.Recordset!KODEOPR
Text2.Enabled = False
End If
If Combo1 = "Admin" Then Cmdhapus.Enabled = False
End Sub



Private Sub COMBO1_Click()
If Combo1.Text = "Kasir" Then
Text1.Text = "KSR" & ADO.Recordset.RecordCount
End If

End Sub

Private Sub Form_Load()
Combo1.AddItem "Kasir"
Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
If Not (KeyAscii >= Asc("a") & Chr(13) And KeyAscii <= Asc("z") & Chr(13) Or (KeyAscii >= Asc("A") & Chr(13) And KeyAscii <= Asc("Z") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeySpace)) Then
KeyAscii = 0
End If
End Sub

