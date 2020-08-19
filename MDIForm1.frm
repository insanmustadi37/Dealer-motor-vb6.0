VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Penjualan Kendaraan Bermotor CV. Medan Jaya Motor Sumatera"
   ClientHeight    =   4755
   ClientLeft      =   4050
   ClientTop       =   3465
   ClientWidth     =   7575
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1560
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Left            =   1920
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1800
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   705
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E075
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E94F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F229
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FB03
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Report 
      Left            =   3060
      Top             =   1980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7699
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnutype 
         Caption         =   "Entry Type Kendaraan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuangsuran 
         Caption         =   "Entry Data Harga Angsuran Kredit"
      End
      Begin VB.Menu mnukendaraan 
         Caption         =   "Entry Data &Kendaraan"
      End
      Begin VB.Menu mnupelanggan 
         Caption         =   "Entry Data &Pelanggan"
      End
      Begin VB.Menu mnutransaksi 
         Caption         =   "Entry Data &Transaksi"
      End
      Begin VB.Menu mnudetail 
         Caption         =   "Entry Detail Angsuran"
      End
      Begin VB.Menu mnugrs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuinfo 
      Caption         =   "&Pencarian"
      Begin VB.Menu mnuinfokendaraan 
         Caption         =   "Informasi Data &Kendaraan"
      End
      Begin VB.Menu mnuinfopelanggan 
         Caption         =   "Informasi Data &Pelanggan"
      End
      Begin VB.Menu mnuinfotransaksi 
         Caption         =   "Informasi Data &Transaksi"
      End
   End
   Begin VB.Menu mnulaporan 
      Caption         =   "&Laporan"
      Begin VB.Menu mnulapkendaraan 
         Caption         =   "Laporan Data &Kendaraan"
      End
      Begin VB.Menu mnulappelanggan 
         Caption         =   "Laporan Data &Pelanggan"
      End
      Begin VB.Menu mnulaptransaksi 
         Caption         =   "Laporan Data &Transaksi"
         Begin VB.Menu mnuselurus 
            Caption         =   "Seluruh Transaksi"
         End
         Begin VB.Menu mnutunai 
            Caption         =   "Penjualan Tunai"
         End
         Begin VB.Menu mnukredit 
            Caption         =   "Penjualan Kredit"
         End
      End
      Begin VB.Menu mnulapdetail 
         Caption         =   "Laporan Detail Angsuran"
      End
      Begin VB.Menu mnulapangsuranpel 
         Caption         =   "Laporan Data Angsuran Pelanggan"
      End
   End
   Begin VB.Menu mnuBantuan 
      Caption         =   "Bantuan"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer
Dim strTemp, LenTemp, n

Private Sub MDIForm_Load()
 strTemp = Me.Caption
    n = 1
    
    Dim Ahari
    Dim SHari As String

    Counter = 0
    Timer2.Interval = 100
    With StatusBar1
        .Panels(1).Width = 4000
        .Panels(1).Alignment = sbrRight
    End With

    Ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
    SHari = Ahari(Abs(Weekday(Date) - 1))
  
        StatusBar1.Panels(2).Text = "" & SHari & ", " & Format(Date, "dd.mm.yyyy")  'Tampilan jam
        StatusBar1.Panels(3).Text = Time                                            'pada status bar
    Timer1.Enabled = True
End Sub

Private Sub mnuBantuan_Click()
MsgBox " Saya Rasa Nggak Perlu Lagi Untuk Dijelaskan Cara Penggunaan Program ini", vbInformation, "Cara Penggunaan"
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(3).Text = Time
End Sub

Private Sub Timer2_Timer()
Dim Kalimat As String
    Dim pnlX1 As Panel
    Set pnlX1 = StatusBar1.Panels(1)
        Kalimat = "CV. Medan Jaya Motor  Berdomisili Di Jalan Glugur By Pass No.91 C Medan - Sumatera Utara, Telp (061)6613670 Fax 66253575"
        Counter = Counter + 1
        DoEvents
        pnlX1.Text = TulisJalan(Counter, Kalimat, 150)
End Sub
Public Function TulisJalan(hitung As Integer, _
                           strKalimat As String, _
                           Panjang As Integer)
    If hitung = Len(strKalimat) + Panjang Then
        hitung = 0
    ElseIf hitung > Len(strKalimat) Then
        TulisJalan = strKalimat & Space(hitung - Len(strKalimat))
    Else
        TulisJalan = Mid(strKalimat, 1, hitung)
    End If
End Function
Private Sub Timer3_Timer()
LenTemp = Len(strTemp)
    Dim Form As String
    LenTemp = Len(strTemp)
    Me.Caption = Left(strTemp, n) + "_"
    n = n + 1
    If n > LenTemp Then
        n = 1
    End If
End Sub

Private Sub mnuangsuran_Click()
Form10.Show
End Sub

Private Sub mnudetail_Click()
Form11.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuinfokendaraan_Click()
Form6.Show
End Sub

Private Sub mnuinfopelanggan_Click()
Form7.Show
End Sub

Private Sub mnuinfotransaksi_Click()
Form8.Show
End Sub

Private Sub mnukendaraan_Click()
Form2.Show
End Sub

Private Sub mnukredit_Click()
Dim fakkredit As String
fakkredit = InputBox("Masukkan Faktur Penjualan Kredit......", "Laporan Penjualan Kredit")
If Not fakkredit = "" Then
    Report.ReportFileName = App.Path & "\rptjualkredit.rpt"
    Report.DataFiles(0) = App.Path & "\penjualan.mdb"
    Report.ReplaceSelectionFormula "{penjualan.faktur} = '" & fakkredit & "' and {penjualan.jns_trans} = 'Kredit'"
    Report.WindowState = crptMaximized
    Report.Action = 7
    Report.Reset
End If
End Sub

Private Sub mnulapangsuranpel_Click()
Form13.Show
End Sub

Private Sub mnulapdetail_Click()
Form12.Show
End Sub

Private Sub mnulapkendaraan_Click()
periode = InputBox("Masukkan Periode.....", "Periode", Format(Date, "dd MMMM yyyy"))
If Not periode = "" Then
    Report.ReportFileName = App.Path & "\rptkendaraan.rpt"
    Report.DataFiles(0) = App.Path & "\penjualan.mdb"
    Report.Formulas(0) = "Periode='" & periode & "'"
    Report.WindowState = crptMaximized
    Report.Action = 7
    Report.Reset
End If
End Sub

Private Sub mnulappelanggan_Click()
Report.ReportFileName = App.Path & "\rptpelanggan.rpt"
Report.DataFiles(0) = App.Path & "\penjualan.mdb"
Report.WindowState = crptMaximized
Report.Action = 7
Report.Reset
End Sub

Private Sub mnupelanggan_Click()
Form1.Show
End Sub

Private Sub mnuselurus_Click()
Report.ReportFileName = App.Path & "\rpttransaksi.rpt"
Report.DataFiles(0) = App.Path & "\penjualan.mdb"
Report.WindowState = crptMaximized
Report.Action = 7
Report.Reset
End Sub

Private Sub mnutransaksi_Click()
Form3.Show
End Sub

Private Sub mnutunai_Click()
Dim faktur As String
faktur = InputBox("Masukkan Faktur Penjualan Tunai...", "Laporan Penjualan Tunai")
If Not faktur = "" Then
    Report.ReportFileName = App.Path & "\rptjualtunai.rpt"
    Report.DataFiles(0) = App.Path & "\penjualan.mdb"
    Report.ReplaceSelectionFormula "{penjualan.faktur}='" & faktur & "' and {Penjualan.Jns_Trans}='Tunai'"
    Report.WindowState = crptMaximized
    Report.Action = 7
    Report.Reset
End If
End Sub

Private Sub mnutype_Click()
Form9.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Form2.Show
    Case 2
        Form1.Show
    Case 3
        Form3.Show
    Case 5
        End
End Select
End Sub
