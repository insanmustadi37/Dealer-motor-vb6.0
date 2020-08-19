VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form LapPembayaran 
   Caption         =   "Laporan Pembayaran Kelompok 6"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4080
      OleObjectBlob   =   "LapPembayaran.frx":0000
      Top             =   2160
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":0234
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin Crystal.CrystalReport CR 
         Left            =   2880
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   2520
         Width           =   1750
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cetak Semua Data"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   1750
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   1560
         Width           =   1750
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1920
         Width           =   1750
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":02A0
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":0308
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":0370
         TabIndex        =   9
         Top             =   2520
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":03E0
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   375
         Left            =   120
         OleObjectBlob   =   "LapPembayaran.frx":045C
         TabIndex        =   11
         Top             =   120
         Width           =   1455
      End
   End
End
Attribute VB_Name = "LapPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo4_Click()
    CR.SelectionFormula = "{BayarCicilan.idkredit}='" & (Combo4) & "'"
    CR.ReportFileName = App.Path & "\lap bayar cicilan per id.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Form_Load()
'On Error Resume Next
Call BukaDB
RSBayarCicilan.Open "Select Distinct TanggalByr From BayarCicilan order By 1", CONN
RSBayarCicilan.Requery
Do Until RSBayarCicilan.EOF
    Combo1.AddItem Format(RSBayarCicilan!TanggalByr, "DD-MMM-YYYY")
    RSBayarCicilan.MoveNext
Loop

Dim RSBulan As New ADODB.Recordset
RSBulan.Open "select distinct month(TanggalByr) as Bulan from BayarCicilan", CONN
Do While Not RSBulan.EOF
    Combo2.AddItem RSBulan!Bulan & Space(5) & MonthName(RSBulan!Bulan)
    RSBulan.MoveNext
Loop

Dim RSTahun As New ADODB.Recordset
RSTahun.Open "select distinct year(TanggalByr)  as Tahun from BayarCicilan", CONN
Do While Not RSTahun.EOF
    Combo3.AddItem RSTahun!Tahun
    RSTahun.MoveNext
Loop
CONN.Close

Call BukaDB
RSBayarCicilan.Open "Select distinct idkredit From BayarCicilan order By 1", CONN
RSBayarCicilan.Requery
Do Until RSBayarCicilan.EOF
    Combo4.AddItem RSBayarCicilan!IdKredit
    RSBayarCicilan.MoveNext
Loop
CONN.Close

Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd

End Sub

Private Sub COMBO1_Click()
    CR.SelectionFormula = "Totext({BayarCicilan.TanggalByr})='" & CDate(Combo1) & "'"
    CR.ReportFileName = App.Path & "\lap bayar cicilan harian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Combo3_Click()
    Call BukaDB
    RSBayarCicilan.Open "select * from BayarCicilan where month(TanggalByr)='" & Val(Left(Combo2, 2)) & "' and year(TanggalByr)='" & (Combo3) & "'", CONN
    If RSBayarCicilan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    CR.SelectionFormula = "Month({BayarCicilan.TanggalByr})=" & Val(Left(Combo2, 2)) & " and Year({BayarCicilan.TanggalByr})=" & Val(Combo3.Text)
    CR.ReportFileName = App.Path & "\LAP bayar cicilan bulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
    CR.Reset
End Sub

Private Sub Command1_Click()
    CR.ReportFileName = App.Path & "\lap bayar cicilan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub
