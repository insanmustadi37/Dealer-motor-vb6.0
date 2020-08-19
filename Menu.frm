VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm Menu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sistem Informasi Penjualan Kendaraan Bermotor CV. Jaya Motor Jakarta"
   ClientHeight    =   8130
   ClientLeft      =   4050
   ClientTop       =   3465
   ClientWidth     =   14940
   LinkTopic       =   "MDIForm1"
   Picture         =   "Menu.frx":0000
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1560
      Top             =   1200
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
            Picture         =   "Menu.frx":16E3A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":16EC7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":16F556
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":16FE30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   7695
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20690
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
   Begin VB.Menu mnfile 
      Caption         =   "&Menu Admin"
      Begin VB.Menu mnoperator 
         Caption         =   "Operator"
      End
      Begin VB.Menu mncustomer 
         Caption         =   "Costumer"
      End
      Begin VB.Menu mnmotor 
         Caption         =   "Motor"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "&Menu Pembayaran"
      Begin VB.Menu mncash 
         Caption         =   "Pembayaran Cash"
      End
      Begin VB.Menu mnkredit 
         Caption         =   "Pembayarab Kredit"
      End
      Begin VB.Menu mnbayarcicilan 
         Caption         =   "Pembayaran Cicilan"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "&Menu Laporan"
      Begin VB.Menu mnlapbeli 
         Caption         =   "Laporan Pembelian"
      End
      Begin VB.Menu mnlapbayar 
         Caption         =   "Laporan Pembayaran"
      End
   End
   Begin VB.Menu logout 
      Caption         =   "&Menu Beralih"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer
Dim strTemp, LenTemp, n

Private Sub Logout_Click()
Login.Show
Menu.Hide
Unload Me
End Sub

Private Sub mnbayarcicilan_Click()
BayarCicilan.Show vbModal
End Sub

Private Sub mncash_Click()
BeliCash.Show vbModal
End Sub

Private Sub mncustomer_Click()
Customer.Show vbModal
End Sub

Private Sub mnkredit_Click()
BeliKredit.Show vbModal
End Sub


Private Sub mnlapbayar_Click()
LapPembayaran.Show vbModal
End Sub

Private Sub mnlapbeli_Click()
LapPembelian.Show vbModal
End Sub

Private Sub mnmotor_Click()
Motor.Show vbModal
End Sub

Private Sub mnoperator_Click()
Operator.Show vbModal
End Sub
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


Private Sub Timer1_Timer()
StatusBar1.Panels(3).Text = Time
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
