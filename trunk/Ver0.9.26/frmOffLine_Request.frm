VERSION 5.00
Begin VB.Form frmOffLine_Request 
   Caption         =   "Device Off Line Request"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4815
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdAPI_OffLine 
      Caption         =   "API OFF LINE"
      Height          =   825
      Left            =   2790
      TabIndex        =   1
      Top             =   420
      Width           =   1545
   End
   Begin VB.CommandButton cmdEQ_OffLine 
      Caption         =   "EQP OFF LINE"
      Height          =   825
      Left            =   540
      TabIndex        =   0
      Top             =   420
      Width           =   1545
   End
End
Attribute VB_Name = "frmOffLine_Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAPI_OffLine_Click()

    Dim intResult               As Integer
    Dim intPortNo               As Integer
    
    Dim strDeviceStatus         As String
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strDeviceStatus)
    
    If intPortNo > 0 Then
        intResult = QUEUE.Put_Send_Command(intPortNo, "QOFA")
    Else
        Call Show_Message("Off Line request fail", "API is not On Line mode.")
        Unload Me
    End If
    
End Sub

Private Sub cmdEQ_OffLine_Click()

    Dim intResult               As Integer
    Dim intPortNo               As Integer
    
    Dim strDeviceStatus         As String
    Dim strDeviceType           As String
    Dim strCommand              As String
    
    strDeviceType = Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5)
    
    Call ENV.Get_Device_Data_by_Name(strDeviceType, intPortNo, strDeviceStatus)
    
    If intPortNo > 0 Then
        If strDeviceType = "CATST" Then
            strCommand = "QOFT"
        ElseIf strDeviceType = "CALOI" Then
            strCommand = "QOFI"
        Else
            strCommand = ""
        End If
        If strCommand <> "" Then
            intResult = QUEUE.Put_Send_Command(intPortNo, strCommand)
        Else
            Call Show_Message("Device name error.", "Device name is not exist.")
        End If
    Else
        Call Show_Message("Off Line rquest fail", "EQP is not On Line mode.")
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    If Get_Device_State(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5)) = cDEVICE_OFFLINE Then
        Me.cmdEQ_OffLine.Enabled = False
    Else
        Me.cmdEQ_OffLine.Enabled = True
    End If
    
    If Get_Device_State("API") = cDEVICE_OFFLINE Then
        Me.cmdAPI_OffLine.Enabled = False
    Else
        Me.cmdAPI_OffLine.Enabled = True
    End If
    
End Sub
