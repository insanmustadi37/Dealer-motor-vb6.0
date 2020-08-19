VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Login 
   Caption         =   "Login Kelompok 6"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "LOGIN"
      Height          =   3015
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7095
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "Login.frx":0000
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6000
         OleObjectBlob   =   "Login.frx":007E
         Top             =   2640
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Terlihat"
         Height          =   435
         Left            =   5520
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   350
         Left            =   2040
         TabIndex        =   0
         Top             =   480
         Width           =   3315
      End
      Begin VB.TextBox Text2 
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "X"
         TabIndex        =   1
         Top             =   1080
         Width           =   3315
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   495
         Left            =   120
         OleObjectBlob   =   "Login.frx":02B2
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   5535
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   9763
      _cy             =   1508
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text2.PasswordChar = ""
    Else
    Text2.PasswordChar = "X"
End If

End Sub

Private Sub Command1_Click()

Call BukaDB


 RSOperator.Open "Select * from Operator where NamaOPR ='" & Text1 & "' and PasswordOPR='" & Text2 & "'", CONN

If Not RSOperator.EOF Then

    If RSOperator!STATUSOPR = "Admin" Then

        Menu.mnfile.Enabled = True

        Menu.mnlaporan.Enabled = True

        Menu.logout.Enabled = True

        Menu.mntransaksi.Enabled = True
        
    Form1.Show
    Login.Hide
    

    ElseIf RSOperator!STATUSOPR = "Kasir" Then

        Menu.mnfile.Visible = False
    
        Menu.mnlaporan.Visible = False
    
        Menu.mntransaksi.Enabled = True

        Menu.logout.Enabled = True

         Form1.Show
        Login.Hide
    

        Call bersih
        Else
        Menu.mnfile.Visible = False
    
        Menu.mnlaporan.Visible = False
    
        Menu.mntransaksi.Enabled = True

        Menu.logout.Enabled = True

        Form1.Show
        Login.Hide
    
    End If
Call bersih
Else

MsgBox "Maaf,Username Atau Password Salah", vbCritical, "Peringatan"
Call bersih
End If
End Sub
  Sub bersih()
    Text1.Text = ""
    Text2.Text = ""
    End Sub

Private Sub Command2_Click()
Dim answer As Integer
 
answer = MsgBox("APAKAH ANDA INGIN KELUAR DARI FORM LOGIN INI????", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbYes Then
  End
End If
End Sub
Private Sub Form_Load()
WindowsMediaPlayer1.URL = App.Path & "\DEWA19 ft NOAH - Semakin Di DepanIklan Yamaha Menembus Langit.mp3"
Call bersih
Skin1.LoadSkin "c:\Program files (x86)\activeskin 4.3\skins\chizh.skn "
Skin1.ApplySkin Me.hWnd


End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

