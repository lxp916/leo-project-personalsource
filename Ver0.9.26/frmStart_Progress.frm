VERSION 5.00
Begin VB.Form frmStart_Progress 
   Caption         =   "JPS starting...."
   ClientHeight    =   1740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   1740
   ScaleWidth      =   9075
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox Picture1 
      Height          =   465
      Left            =   90
      ScaleHeight     =   405
      ScaleWidth      =   8805
      TabIndex        =   0
      Top             =   1140
      Width           =   8865
      Begin VB.Shape shpProgress 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   435
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Label lblModule_Name 
      AutoSize        =   -1  'True
      Caption         =   "Main function"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Processing module"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2910
   End
End
Attribute VB_Name = "frmStart_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Private Const HWND_TOPMOST = -1
  Private Const SWP_NOMOVE = &H2
  Private Const SWP_NOSIZE = &H1

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    Me.shpProgress.Width = 15
    
End Sub
