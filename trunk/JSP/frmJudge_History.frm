VERSION 5.00
Begin VB.Form frmJudge_History 
   Caption         =   "Judge History"
   ClientHeight    =   11130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   11130
   ScaleWidth      =   20370
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      Height          =   10425
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   20355
      Begin VB.TextBox txtJudge_History 
         Height          =   10095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   3
         Top             =   210
         Width           =   20115
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   10650
      TabIndex        =   1
      Top             =   10530
      Width           =   1245
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   525
      Left            =   8520
      TabIndex        =   0
      Top             =   10530
      Width           =   1245
   End
End
Attribute VB_Name = "frmJudge_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub
