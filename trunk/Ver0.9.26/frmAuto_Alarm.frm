VERSION 5.00
Begin VB.Form frmAuto_Alarm 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Defect Occurred"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   19800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   19800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tmrIndicator 
      Interval        =   500
      Left            =   360
      Top             =   900
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "CONFIRM"
      Default         =   -1  'True
      Height          =   615
      Left            =   8760
      TabIndex        =   0
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   300
      Width           =   120
   End
End
Attribute VB_Name = "frmAuto_Alarm"
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
    Me.lblTitle.BackColor = vbRed
    Me.lblTitle.ForeColor = vbBlack
    Me.BackColor = vbRed
    
End Sub

Private Sub OKButton_Click()

    Unload Me
    
End Sub

Private Sub tmrIndicator_Timer()

    If Me.lblTitle.BackColor = vbRed Then
        Me.lblTitle.BackColor = vbBlack
        Me.lblTitle.ForeColor = vbYellow
        Me.BackColor = vbBlack
    Else
        Me.lblTitle.BackColor = vbRed
        Me.lblTitle.ForeColor = vbBlack
        Me.BackColor = vbRed
    End If
    
End Sub
