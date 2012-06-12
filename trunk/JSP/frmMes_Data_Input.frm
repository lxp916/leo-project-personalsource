VERSION 5.00
Begin VB.Form frmMes_Data_Input 
   Caption         =   "Data Input"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3945
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   3945
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.CommandButton cmdMes_Input 
      Caption         =   "Input Data"
      Height          =   525
      Left            =   1440
      TabIndex        =   2
      Top             =   870
      Width           =   1245
   End
   Begin VB.TextBox txtMes_Data 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1770
      TabIndex        =   1
      Top             =   330
      Width           =   2085
   End
   Begin VB.Label lblRow 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   75
      TabIndex        =   0
      Top             =   420
      Width           =   1635
   End
End
Attribute VB_Name = "frmMes_Data_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdMes_Input_Click()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strCommand                  As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intRow                      As Integer
    Dim intSpace                    As Integer
    Dim intSize                     As Integer
    
    intRow = CInt(Me.lblRow.Caption)
    frmMain.flxMES_Data.TextMatrix(intRow, 1) = Me.txtMes_Data.Text
    
    Select Case intRow
    Case 1:
        pubCST_INFO.CSTID = Me.txtMes_Data.Text
    Case 2:
        pubCST_INFO.OWNER = Me.txtMes_Data.Text
    Case 3:
        pubCST_INFO.PROCESS_NUM = Me.txtMes_Data.Text
    Case 17:
        pubPANEL_INFO.PANELID = Me.txtMes_Data.Text
    End Select
    
    strPath = App.PATH & "\Env\"
    strFileName = "RBBC.txt"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        Line Input #intFileNum, strCommand
        
        Close intFileNum
        
        strTemp = Left(strCommand, 5)
        strCommand = Mid(strCommand, 6)
        
        intSpace = cSIZE_PANELID - Len(frmMain.flxMES_Data.TextMatrix(17, 1))
        strTemp = strTemp & frmMain.flxMES_Data.TextMatrix(17, 1) & Space(intSpace)
        strCommand = Mid(strCommand, cSIZE_PANELID + 1)
        
        intSize = 3 + cSIZE_INFO_LENGTH + cSIZE_CSTID_MES
        strTemp = strTemp & Left(strCommand, intSize)
        strCommand = Mid(strCommand, intSize + 1)
        
        intSpace = cSIZE_PFCD - Len(frmMain.flxMES_Data.TextMatrix(1, 1))
        strTemp = strTemp & frmMain.flxMES_Data.TextMatrix(1, 1) & Space(intSpace)
        strCommand = Mid(strCommand, cSIZE_PFCD + 1)
        
        intSpace = cSIZE_OWNER_MES - Len(frmMain.flxMES_Data.TextMatrix(2, 1))
        strTemp = strTemp & frmMain.flxMES_Data.TextMatrix(2, 1) & Space(intSpace)
        strCommand = Mid(strCommand, cSIZE_OWNER_MES + 1)
        
        intSpace = cSIZE_PROCESSNUM_MES - Len(frmMain.flxMES_Data.TextMatrix(3, 1))
        strTemp = strTemp & frmMain.flxMES_Data.TextMatrix(3, 1) & Space(intSpace)
        strCommand = Mid(strCommand, cSIZE_PROCESSNUM_MES + 1)
        
        intSize = cSIZE_PORTID_MES + cSIZE_PORTTYPE_MES + cSIZE_DESTFAB_MES + cSIZE_PANELCOUNT_MES
        intSize = intSize + cSIZE_RMANO_MES + cSIZE_OQCNO_MES + cSIZE_SOURCE_FAB_MES
        intSize = intSize + (cSIZE_CST_SPARE_MES * 5) + cSIZE_INFO_LENGTH + cSIZE_SLOTNO_MES
        strTemp = strTemp & Left(strCommand, intSize)
        strCommand = Mid(strCommand, intSize + 1)
        
        intSpace = cSIZE_PANELID - Len(frmMain.flxMES_Data.TextMatrix(17, 1))
        strTemp = strTemp & frmMain.flxMES_Data.TextMatrix(17, 1) & Space(intSpace)
        strCommand = Mid(strCommand, cSIZE_PANELID + 1)
        
        strTemp = strTemp & strCommand
        
        intFileNum = FreeFile
        
        Open strPath & strFileName For Output As intFileNum
        
        Print #intFileNum, strTemp
        
        Close intFileNum
        
        strCommand = Left(strTemp, 1) & "RABC" & Mid(strTemp, 6)
        strFileName = "RABC.txt"
        intFileNum = FreeFile
        
        Open strPath & strFileName For Output As intFileNum
        
        Print #intFileNum, strCommand
        
        Close intFileNum
    End If
    
    Unload Me
    
End Sub
