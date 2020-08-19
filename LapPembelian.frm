VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form LapPembelian 
   Caption         =   "Laporan Pembelian Cash Dan Kredit Kelompok 6"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7320
      OleObjectBlob   =   "LapPembelian.frx":0000
      Top             =   1920
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Pembelian Kredit"
      Height          =   3615
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Cetak Semua Data"
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   1750
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   1750
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "LapPembelian.frx":0234
         TabIndex        =   13
         Top             =   600
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "LapPembelian.frx":02A0
         TabIndex        =   14
         Top             =   1560
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "LapPembelian.frx":0308
         TabIndex        =   15
         Top             =   1920
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":0370
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":03EA
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Pembelian Cash"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":0466
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak Semua Data"
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2775
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   1750
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1560
         Width           =   1750
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   240
         OleObjectBlob   =   "LapPembelian.frx":04E0
         TabIndex        =   10
         Top             =   600
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":054C
         TabIndex        =   11
         Top             =   1560
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":05B4
         TabIndex        =   12
         Top             =   1920
         Width           =   1005
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "LapPembelian.frx":061C
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
   End
End
Attribute VB_Name = "LapPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
'cari data tanggal di tabel belicash
RSBeliCash.Open "Select Distinct Tanggal From BeliCash order By 1", CONN
RSBeliCash.Requery
Do Until RSBeliCash.EOF
    'tampilkan dalam combo1
    Combo1.AddItem Format(RSBeliCash!Tanggal, "DD-MMM-YYYY")
    RSBeliCash.MoveNext
Loop

Dim RSBulan As New ADODB.Recordset
'cari bulan dalam tabel belicash
RSBulan.Open "select distinct month(Tanggal) as Bulan from BeliCash", CONN
Do While Not RSBulan.EOF
    'tampilkan dalam combo2
    Combo2.AddItem RSBulan!Bulan & Space(5) & MonthName(RSBulan!Bulan)
    RSBulan.MoveNext
Loop

Dim RSTahun As New ADODB.Recordset
'cari tahun di tabel belicash
RSTahun.Open "select distinct year(Tanggal)  as Tahun from BeliCash", CONN
Do While Not RSTahun.EOF
    'tampilkan dalam combo3
    Combo3.AddItem RSTahun!Tahun
    RSTahun.MoveNext
Loop


RSBeliKredit.Open "Select Distinct Tanggal From BeliKredit order By 1", CONN
RSBeliKredit.Requery
Do Until RSBeliKredit.EOF
    Combo4.AddItem Format(RSBeliKredit!Tanggal, "DD-MMM-YYYY")
    RSBeliKredit.MoveNext
Loop

Dim RSBulanKredit As New ADODB.Recordset
RSBulanKredit.Open "select distinct month(Tanggal) as Bulan from BeliKredit", CONN
Do While Not RSBulanKredit.EOF
    Combo5.AddItem RSBulanKredit!Bulan & Space(5) & MonthName(RSBulanKredit!Bulan)
    RSBulanKredit.MoveNext
Loop

Dim RSTahunKredit As New ADODB.Recordset
RSTahunKredit.Open "select distinct year(Tanggal)  as Tahun from BeliKredit", CONN
Do While Not RSTahunKredit.EOF
    Combo6.AddItem RSTahunKredit!Tahun
    RSTahunKredit.MoveNext
Loop

CONN.Close

Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub COMBO1_Click()
    CrystalReport1.SelectionFormula = "Totext({BeliCash.Tanggal})='" & CDate(Combo1) & "'"
    CrystalReport1.ReportFileName = App.Path & "\lap beli cash harian.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Combo3_Click()
    Call BukaDB
    RSBeliCash.Open "select * from BeliCash where month(Tanggal)='" & Val(Left(Combo2, 2)) & "' and year(Tanggal)='" & (Combo3) & "'", CONN
    If RSBeliCash.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CrystalReport1.SelectionFormula = "Month({BeliCash.Tanggal})=" & Val(Left(Combo2, 2)) & " and Year({BeliCash.Tanggal})=" & Val(Combo3.Text)
    CrystalReport1.ReportFileName = App.Path & "\LAP beli cash bulanan.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Combo4_Click()
    CrystalReport1.SelectionFormula = "Totext({BeliKredit.Tanggal})='" & CDate(Combo4) & "'"
    CrystalReport1.ReportFileName = App.Path & "\LAP BELI KREDIT HARIAN.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Combo6_Click()
Call BukaDB
RSBeliKredit.Open "select * from BeliKredit where month(Tanggal)='" & Val(Left(Combo5, 2)) & "' and year(Tanggal)='" & (Combo6) & "'", CONN
If RSBeliKredit.EOF Then
    MsgBox "Data tidak ditemukan"
    Exit Sub
    Combo4.SetFocus
End If
CrystalReport1.SelectionFormula = "Month({BeliKredit.Tanggal})=" & Val(Left(Combo5, 2)) & " and Year({BeliKredit.Tanggal})=" & Val(Combo6.Text)
CrystalReport1.ReportFileName = App.Path & "\LAP BELI KREDIT BULANAN.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1

End Sub


Private Sub Command1_Click()
    CrystalReport1.ReportFileName = App.Path & "\lap beli cash.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub Command2_Click()
    CrystalReport1.ReportFileName = App.Path & "\lap beli kredit.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub
