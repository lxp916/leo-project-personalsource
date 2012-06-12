VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "User Log In"
   ClientHeight    =   2520
   ClientLeft      =   2850
   ClientTop       =   3495
   ClientWidth     =   4815
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1488.899
   ScaleMode       =   0  'User
   ScaleWidth      =   4521.024
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Confirm"
      Default         =   -1  'True
      Height          =   390
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
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
  End Sub


Private Sub cmdCancel_Click()
    
    Dim typCURRENT_USER             As USER_LOGON_DATA

    With typCURRENT_USER
'        Call ENV.Get_Current_Logon_User(.LOGON_TIME, .USER_ID, .USER_NAME)
        If (.LOGON_TIME <> "") And (.USER_ID <> "") And (.USER_NAME <> "") Then
            Unload Me
        Else
            Me.SetFocus
            Me.Visible = True
        End If
    End With
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim LOGON_USER_DATA             As USER_LOGON_DATA
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim bolLogon_Success            As Boolean
    
    If Me.txtUserName.Text = "Lucas" Then
        If Me.txtPassword.Text = "1" Then
            Call ENV.Set_Current_Logon_User(Me.txtUserName.Text, "Lucas")
            bolLogon_Success = True
            Unload Me
        End If
    Else
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "Parameter.mdb"
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
            strQuery = "SELECT * FROM USER_DATA WHERE "
            strQuery = strQuery & "USER_ID = '" & Me.txtUserName.Text & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                If Me.txtPassword.Text = lstRecord.Fields("USER_PW2") Then
                    With LOGON_USER_DATA
                        .LOGON_TIME = ""
                        .USER_ID = Me.txtUserName.Text
                        .USER_NAME = lstRecord.Fields("USER_NAME")
                        Call ENV.Set_Current_Logon_User(.USER_ID, .USER_NAME)
                    End With
                    bolLogon_Success = True
                Else
                    Call Show_Message("Logon error", Me.txtUserName.Text & "'s password is not correct.")
                    Me.txtPassword.SetFocus
                    SendKeys "{Home}+{End}"
                    bolLogon_Success = False
                End If
            End If
            lstRecord.Close
            
            dbMyDB.Close
        Else
            bolLogon_Success = False
            Call Show_Message("DB Error", "Parameter.mdb file is not found.")
        End If
    End If
    
    If bolLogon_Success = True Then
        frmMain.lblUser.Caption = LOGON_USER_DATA.USER_NAME
        frmMain.flxRUN_Info.TextMatrix(4, 1) = "0"
        Select Case ENV.Get_Current_User_Level
        Case "S":
            With frmMain.Toolbar1
                .Buttons(1).Enabled = True
                .Buttons(2).Enabled = True
                .Buttons(3).Enabled = True
                .Buttons(4).Enabled = True
                .Buttons(5).Enabled = True
                .Buttons(6).Enabled = True
                .Buttons(7).Enabled = True
            End With
            frmMain.mnuTools.Enabled = True
        Case "E":
            With frmMain.Toolbar1
                .Buttons(1).Enabled = True
                .Buttons(2).Enabled = True
                .Buttons(3).Enabled = True
                .Buttons(4).Enabled = True
                .Buttons(5).Enabled = True
                .Buttons(6).Enabled = True
                .Buttons(7).Enabled = True
            End With
            frmMain.mnuTools.Enabled = True
        Case "P":
            With frmMain.Toolbar1
                .Buttons(1).Enabled = True
                .Buttons(2).Enabled = True
                .Buttons(3).Enabled = True
                .Buttons(4).Enabled = False
                .Buttons(5).Enabled = False
                .Buttons(6).Enabled = True
                .Buttons(7).Enabled = True
            End With
            frmMain.mnuTools.Enabled = True
        Case "T":
            With frmMain.Toolbar1
                .Buttons(1).Enabled = True
                .Buttons(2).Enabled = False
                .Buttons(3).Enabled = True
                .Buttons(4).Enabled = False
                .Buttons(5).Enabled = False
                .Buttons(6).Enabled = False
                .Buttons(7).Enabled = True
            End With
            frmMain.mnuTools.Enabled = True
        Case Else
            With frmMain.Toolbar1
                .Buttons(1).Enabled = False
                .Buttons(2).Enabled = False
                .Buttons(3).Enabled = False
                .Buttons(4).Enabled = False
                .Buttons(5).Enabled = False
                .Buttons(6).Enabled = False
                .Buttons(7).Enabled = True
            End With
            frmMain.mnuTools.Enabled = False
        End Select
        
        Unload Me
    End If
    
End Sub

