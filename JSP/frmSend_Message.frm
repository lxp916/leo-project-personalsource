VERSION 5.00
Begin VB.Form frmSend_Message 
   Caption         =   "JPS Message"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   9195
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   525
      Left            =   4590
      TabIndex        =   3
      Top             =   1200
      Width           =   1245
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   525
      Left            =   3150
      TabIndex        =   2
      Top             =   1200
      Width           =   1245
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      MaxLength       =   100
      TabIndex        =   1
      Top             =   480
      Width           =   8655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Send Message to EQ"
      BeginProperty Font 
         Name            =   "Gulim"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   270
      TabIndex        =   0
      Top             =   150
      Width           =   2610
   End
End
Attribute VB_Name = "frmSend_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSend_Click()

    Dim intPortNo           As Integer
    
    Dim strStatus           As String
    
    If Me.txtMessage.Text <> "" Then
        Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
        
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBAM" & Me.txtMessage.Text)
        Else
            Call SaveLog("mnuSend_Click", "Port Number : 0")
        End If
    End If
    
    Unload Me
    
End Sub
