VERSION 5.00
Begin VB.Form frmSystem_Log 
   Caption         =   "System Log"
   ClientHeight    =   10845
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Frame fmeLog 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20355
      Begin VB.ListBox lstSystem_Log 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   10560
         ItemData        =   "frmSystem_Log.frx":0000
         Left            =   60
         List            =   "frmSystem_Log.frx":0002
         TabIndex        =   1
         Top             =   180
         Width           =   20235
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmSystem_Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Left = 0
    Me.Top = 0
    Me.Height = 11535
    Me.Width = 20490

End Sub

Private Sub lstSystem_Log_DblClick()

    Dim strFileName                 As String
    Dim strPath                     As String
    
    Dim varResult                   As Variant
    
    strPath = App.PATH & "\Log\"
    strFileName = Format(DATE, "YYYYMMDD") & "_" & Format(TIME, "HH") & ".Log"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        varResult = Shell("NotePad.exe " & strPath & strFileName, vbNormalFocus)
    Else
        strFileName = Format(DATE - (1 / 24), "YYYYMMDD") & "_" & Format(TIME - (1 / 24), "HH") & ".Log"
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            varResult = Shell("NotePad.exe " & strPath & strFileName, vbNormalFocus)
        End If
    End If

End Sub

Private Sub mnuExit_Click()

    Me.Hide
    
End Sub
