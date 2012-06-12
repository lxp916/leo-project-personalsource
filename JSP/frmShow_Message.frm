VERSION 5.00
Begin VB.Form frmShow_Message 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tmrAuto_Close 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11430
      Top             =   1080
   End
   Begin VB.Timer tmrBackColor 
      Interval        =   500
      Left            =   11910
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12375
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         Caption         =   "TEST"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmShow_Message"
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
    Me.Frame1.BackColor = vbRed
    Me.lblMessage.BackColor = vbRed
    Me.lblMessage.ForeColor = vbBlack
    
End Sub

Private Sub OKButton_Click()

    Unload Me
    
End Sub

Private Sub tmrBackColor_Timer()

    If Me.Frame1.BackColor = vbRed Then
        Me.Frame1.BackColor = vbBlack
        Me.lblMessage.BackColor = vbBlack
        Me.lblMessage.ForeColor = vbYellow
    Else
        Me.Frame1.BackColor = vbRed
        Me.lblMessage.BackColor = vbRed
        Me.lblMessage.ForeColor = vbBlack
    End If
    
End Sub
