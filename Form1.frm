VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9045
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   13155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   8175
      Left            =   -480
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   8115
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   0
      Width           =   6615
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   4200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9840
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      TabIndex        =   3
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "MENUJU FORM MENU MOHON BERSABAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Label2.Caption = Label2.Caption + 1
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then
MsgBox "PROSES SELESAI", vbInformation
Unload Me
Menu.Show
Login.Hide
End If
End Sub

