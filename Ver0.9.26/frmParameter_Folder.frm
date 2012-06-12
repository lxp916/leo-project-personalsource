VERSION 5.00
Begin VB.Form frmParameter_Folder 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "FOLER SELECT"
   ClientHeight    =   3705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   135
      Width           =   2580
   End
   Begin VB.DirListBox dirPath 
      Height          =   2610
      Left            =   135
      TabIndex        =   3
      Top             =   540
      Width           =   2580
   End
   Begin VB.DriveListBox drvPath 
      Height          =   300
      Left            =   135
      TabIndex        =   2
      Top             =   3240
      Width           =   2580
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   2790
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2790
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmParameter_Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub dirPath_Change()

    Me.txtPath.Text = Me.dirPath.PATH
    
End Sub

Private Sub drvPath_Change()

    Me.dirPath.PATH = Me.drvPath.Drive
    
End Sub

Private Sub OKButton_Click()

    Unload Me
    
End Sub
