VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVersion 
   Caption         =   "Version"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   11025
   ScaleWidth      =   20370
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.FileListBox fleFile 
      Height          =   450
      Left            =   16230
      TabIndex        =   4
      Top             =   10530
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   10425
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   20355
      Begin MSFlexGridLib.MSFlexGrid flxVersion 
         Height          =   10155
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   20175
         _ExtentX        =   35586
         _ExtentY        =   17912
         _Version        =   393216
         Rows            =   1
         Cols            =   9
      End
   End
   Begin VB.CommandButton cmdGet_Version 
      Caption         =   "Get Version"
      Height          =   525
      Left            =   8520
      TabIndex        =   1
      Top             =   10530
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   10650
      TabIndex        =   0
      Top             =   10530
      Width           =   1245
   End
End
Attribute VB_Name = "frmVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdGet_Version_Click()

    Dim FTP_OBJ                         As New clsFTP
    
    Dim typVERSION_DATA                 As VERSION_DATA
    
    Dim strVersion_FileName()           As String
    
    Dim strRemote_Path                  As String
    Dim strLocal_Path                   As String
    Dim strFileName                     As String
    
    Dim intFileNum                      As Integer
    Dim intFile_Count                   As Integer
    Dim intIndex                        As Integer
    
    Me.flxVersion.Rows = 1
    
    'Download version files from DFS
    strLocal_Path = App.PATH & "\Env\"
    If FTP_OBJ.Init_FTP_Client = True Then
        Call FTP_OBJ.Open_Session
        strRemote_Path = FTP_OBJ.Get_Path(cFTP_DEFECT)
        If Right(strRemote_Path, 1) <> "\" Then
            strRemote_Path = strRemote_Path & "\" & "EQ_Config\" & "JPS\" & "Version"
        End If
        strFileName = FTP_OBJ.FTP_Get_FileList("Version_*.dat", strRemote_Path)
        If FTP_OBJ.FTP_Get_File_from_List(strRemote_Path, strLocal_Path, strLocal_Path, strFileName) = True Then
            Call SaveLog("cmdGet_Version_Click", "Version files download complete.")
        Else
            Call SaveLog("cmdGet_Version_Click", "Version files download fail.")
        End If
        FTP_OBJ.Close_Session
        FTP_OBJ.Disconnect_FTP_Client
    End If
    
    Me.fleFile.PATH = strLocal_Path
    Me.fleFile.Pattern = "VERSION_*.dat"
    
    If Me.fleFile.ListCount > 0 Then
        For intIndex = 0 To Me.fleFile.ListCount - 1
            Call Get_Version_Data_by_FileName(typVERSION_DATA, strLocal_Path & Me.fleFile.List(intIndex))
            If typVERSION_DATA.EQ_VERSION <> "" Then
                Call Add_Row(typVERSION_DATA)
            End If
        Next intIndex
    End If
    
End Sub

Private Sub Form_Load()

    Call Init_Grid
    
    Me.Left = 0
    Me.Top = 0
    Me.Height = 11535
    Me.Width = 20490

End Sub

Private Sub Init_Grid()

    Dim intRow                      As Integer
    Dim intCol                      As Integer
    
    With Me.flxVersion
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(3) = 1500
        .ColWidth(4) = 1700
        .ColWidth(5) = 1700
        .ColWidth(6) = 3200
        .ColWidth(7) = 3200
        .ColWidth(8) = 3200
        
        .TextMatrix(0, 0) = "Machine ID"
        .TextMatrix(0, 1) = "JPS VERSION"
        .TextMatrix(0, 2) = "EQ VERSION"
        .TextMatrix(0, 3) = "JPS Name"
        .TextMatrix(0, 4) = "Installday"
        .TextMatrix(0, 5) = "USER"
        .TextMatrix(0, 6) = "JPS Set-Up Path"
        .TextMatrix(0, 7) = "JPS Log Path"
        .TextMatrix(0, 8) = "JPS Server Path"
    End With

End Sub

Private Sub Add_Row(pVERSION_DATA As VERSION_DATA)

    Dim intRow      As Integer
    Dim intCol      As Integer
    
    With Me.flxVersion
        intRow = .Rows
        .AddItem pVERSION_DATA.MACHINE_ID
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 2
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        .Col = .Cols - 1
        .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(intRow, 1) = pVERSION_DATA.JPS_VERSION
        .TextMatrix(intRow, 2) = pVERSION_DATA.EQ_VERSION
        .TextMatrix(intRow, 3) = pVERSION_DATA.JPS_NAME
        .TextMatrix(intRow, 4) = pVERSION_DATA.INSTALL_DAY
        .TextMatrix(intRow, 5) = pVERSION_DATA.USER
        .TextMatrix(intRow, 6) = pVERSION_DATA.JPS_SETUP_PATH
        .TextMatrix(intRow, 7) = pVERSION_DATA.JPS_LOG_PATH
        .TextMatrix(intRow, 8) = pVERSION_DATA.JPS_SERVER_PATH
    End With
    
End Sub
