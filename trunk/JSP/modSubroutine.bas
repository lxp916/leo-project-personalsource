Attribute VB_Name = "modSubroutine"
Option Explicit

'==========================================================================================================
'
'  Modify Date : 2011. 12. 19
'  Modify by K.H. KIM
'  Content
'    - Before : This function write log data to log file.
'    - After : Put log data to log queue memory after make up log data.
'
'
'==========================================================================================================

Public Sub SaveLog(ByVal Routine_name As String, ByVal msg As String)

    Dim strTime                         As String
    Dim strDate                         As String
    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strLog                          As String
    
    Dim intResult                       As Integer
    
    strTime = Format(TIME, "HH:MM:SS")
    strDate = Format(DATE, "YYYY-MM-DD")
    strFileName = Format(DATE, "YYYYMMDD") & "_" & Format(TIME, "HH") & ".Log"
    strPath = App.PATH & "\Log\"
    strLog = "[" & Routine_name & " " & strDate & " " & strTime & "]  " & msg
    
    intResult = QUEUE.Put_Log_Data(strPath & "," & strFileName & "," & strLog)
    
    Call Add_System_Log(strLog)

End Sub

'==========================================================================================================
'
'  Modify Date : 2011. 12. 19
'  Modify by K.H. KIM
'  Content
'    - Get log data from queue memory append the data to log file
'
'
'==========================================================================================================

Public Sub Write_Log(ByVal pLog As String)

    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strLog                          As String
    
    Dim intPos                          As Integer
    Dim intFileNum                      As Integer
    
    intPos = InStr(pLog, ",")
    If intPos > 0 Then
        strPath = Left(pLog, intPos - 1)
        pLog = Mid(pLog, intPos + 1)
        
        intPos = InStr(pLog, ",")
        strFileName = Left(pLog, intPos - 1)
        strLog = Mid(pLog, intPos + 1)
        
        intFileNum = FreeFile
        
        Open strPath & strFileName For Append As intFileNum
        
        Print #intFileNum, strLog
        
        Close intFileNum
    End If
    
End Sub

Public Sub Add_System_Log(ByVal pLog As String)

    Dim intDelete_Count                 As Integer
    Dim intIndex                        As Integer
    
    intDelete_Count = frmSystem_Log.lstSystem_Log.ListCount - 150
    
    If 0 < intDelete_Count Then
        For intIndex = 1 To intDelete_Count
            frmSystem_Log.lstSystem_Log.RemoveItem (0)
        Next intIndex
    End If
    frmSystem_Log.lstSystem_Log.AddItem pLog
    frmSystem_Log.lstSystem_Log.Selected(frmSystem_Log.lstSystem_Log.ListCount - 1) = True
    
End Sub

Public Sub Show_Message(ByVal pCaption As String, ByVal pMESSAGE As String)

    Load frmShow_Message
    frmShow_Message.Caption = pCaption
    frmShow_Message.lblMessage.Caption = pMESSAGE
    frmShow_Message.Show
    
End Sub

Public Sub State_Change(ByVal pPortID As Integer, ByVal pDevice As String, ByVal pState As String)

    Dim intRow              As Integer
    
    Select Case pDevice
    Case "API":
        intRow = 2
    Case "PG":
        intRow = 3
    Case "CATST":
        intRow = 1
    Case "CALOI":
        intRow = 1
    End Select
    
    If intRow > 0 Then
        With frmMain.flxStatus
            .TextMatrix(intRow, 1) = pState
            .Row = intRow
            .Col = 1
            If pState = cDEVICE_ONLINE Then
                .CellBackColor = vbGreen
                .CellForeColor = vbBlack
            Else
                .CellBackColor = vbRed
                .CellForeColor = vbBlack
            End If
        End With
        
        Call ENV.Set_Device_Info_by_Index(pPortID, pDevice, pState)
    End If
    
End Sub

Public Function Get_Device_State(ByVal pDevice As String) As String

    Dim intRow              As Integer
    
    Select Case pDevice
    Case "PROBER":
        intRow = 1
    Case "API":
        intRow = 2
    Case "PG":
        intRow = 3
    End Select
    
    Get_Device_State = frmMain.flxStatus.TextMatrix(intRow, 1)

End Function

Public Sub Decode_CST_Information_Elements(ByVal pCommand As String, pCST_INFO_ELEMENTS As CST_INFO_ELEMENTS)

    Dim intIndex                    As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    With pCST_INFO_ELEMENTS
        .CSTID = Mid(pCommand, 1, cSIZE_CSTID_MES)
        frmMain.flxMES_Data.TextMatrix(0, 1) = .CSTID
        pCommand = Mid(pCommand, cSIZE_CSTID_MES + 1)
        
        .PFCD = Mid(pCommand, 1, cSIZE_PFCD)
        frmMain.flxMES_Data.TextMatrix(1, 1) = .PFCD
        pCommand = Mid(pCommand, cSIZE_PFCD + 1)
        frmMain.flxEQ_Information.TextMatrix(2, 1) = .PFCD
        frmMain.flxAPI_Information.TextMatrix(2, 1) = .PFCD
        
        .OWNER = Mid(pCommand, 1, cSIZE_OWNER_MES)
        frmMain.flxMES_Data.TextMatrix(2, 1) = .OWNER
        pCommand = Mid(pCommand, cSIZE_OWNER_MES + 1)
        
        .PROCESS_NUM = Mid(pCommand, 1, cSIZE_PROCESSNUM_MES)
        frmMain.flxMES_Data.TextMatrix(3, 1) = .PROCESS_NUM
        pCommand = Mid(pCommand, cSIZE_PROCESSNUM_MES + 1)
        
        .PORTID = Mid(pCommand, 1, cSIZE_PORTID_MES)
        frmMain.flxMES_Data.TextMatrix(4, 1) = .PORTID
        pCommand = Mid(pCommand, cSIZE_PORTID_MES + 1)
        
        .PORT_TYPE = Mid(pCommand, 1, cSIZE_PORTTYPE_MES)
        frmMain.flxMES_Data.TextMatrix(5, 1) = .PORT_TYPE
        pCommand = Mid(pCommand, cSIZE_PORTTYPE_MES + 1)
        
        .DESTINATION_FAB = Mid(pCommand, 1, cSIZE_DESTFAB_MES)
        frmMain.flxMES_Data.TextMatrix(6, 1) = .DESTINATION_FAB
        pCommand = Mid(pCommand, cSIZE_DESTFAB_MES + 1)
        
        .PANEL_COUNT = Mid(pCommand, 1, cSIZE_PANELCOUNT_MES)
        frmMain.flxMES_Data.TextMatrix(7, 1) = .PANEL_COUNT
        pCommand = Mid(pCommand, cSIZE_PANELCOUNT_MES + 1)
        
        .RMANO = Mid(pCommand, 1, cSIZE_RMANO_MES)
        frmMain.flxMES_Data.TextMatrix(8, 1) = .RMANO
        pCommand = Mid(pCommand, cSIZE_RMANO_MES + 1)
        
        .OQCNO = Mid(pCommand, 1, cSIZE_OQCNO_MES)
        frmMain.flxMES_Data.TextMatrix(9, 1) = .OQCNO
        pCommand = Mid(pCommand, cSIZE_OQCNO_MES + 1)
        
        .SOURCE_FAB = Mid(pCommand, 1, cSIZE_SOURCE_FAB_MES)
        frmMain.flxMES_Data.TextMatrix(10, 1) = .SOURCE_FAB
        pCommand = Mid(pCommand, cSIZE_SOURCE_FAB_MES + 1)
        
        If Len(pCommand) > 25 Then
            For intIndex = 1 To 4
                .CST_SPARE(intIndex) = Mid(pCommand, 1, cSIZE_CST_SPARE_MES)
                frmMain.flxMES_Data.TextMatrix(10 + intIndex, 1) = .CST_SPARE(intIndex)
                pCommand = Mid(pCommand, cSIZE_CST_SPARE_MES)
            Next intIndex
    
            .CST_SPARE(5) = Mid(pCommand, cSIZE_CST_SPARE_MES)
            frmMain.flxMES_Data.TextMatrix(15, 1) = .CST_SPARE(5)
        End If
    End With
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Decode_CST_Information_Elements", ErrMsg)
    Call Show_Message("RUN Time error", ErrMsg)
    
End Sub

Public Sub Decode_PANEL_Information_Elements(ByVal pCommand As String, pPANEL_INFO_ELEMENTS As PANEL_INFO_ELEMENTS, ByVal pPFCD As String)

    Dim typPFCD_PID                     As PFCD_DATA
    
    Dim strRemote_Path                  As String
    Dim strLocal_Path                   As String
    Dim strFileName                     As String
    
    Dim intIndex                        As Integer
    
    With pPANEL_INFO_ELEMENTS
        .SLOT_NUM = Mid(pCommand, 1, cSIZE_SLOTNO_MES)
        frmMain.flxMES_Data.TextMatrix(16, 1) = .SLOT_NUM
        pCommand = Mid(pCommand, cSIZE_SLOTNO_MES + 1)
        
        .PANELID = Mid(pCommand, 1, cSIZE_PANELID)
        frmMain.flxMES_Data.TextMatrix(17, 1) = .PANELID
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        
        .LIGHT_ON_PANEL_GRADE = Mid(pCommand, 1, cSIZE_LIGHTON_PNL_GRADE_MES)
        frmMain.flxMES_Data.TextMatrix(18, 1) = .LIGHT_ON_PANEL_GRADE
        pCommand = Mid(pCommand, cSIZE_LIGHTON_PNL_GRADE_MES + 1)
        
        .LIGHT_ON_REASON_CODE = Mid(pCommand, 1, cSIZE_LIGHTON_REASON_MES)
        frmMain.flxMES_Data.TextMatrix(19, 1) = .LIGHT_ON_REASON_CODE
        pCommand = Mid(pCommand, cSIZE_LIGHTON_REASON_MES + 1)
        
'        .CELL_LINE_RESCUE_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
'        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
'        .CELL_REPAIR_JUDGE_GRADE = Mid(pCommand, 1, cSIZE_CELLREPAIR_GRADE_MES)
'        pCommand = Mid(pCommand, cSIZE_CELLREPAIR_GRADE_MES + 1)
        
        .TFT_REPAIR_GRADE = Mid(pCommand, 1, cSIZE_TFT_REPAIR_GRADE_MES)
        frmMain.flxMES_Data.TextMatrix(22, 1) = .TFT_REPAIR_GRADE
        pCommand = Mid(pCommand, cSIZE_TFT_REPAIR_GRADE_MES + 1)
        
        .CF_PANELID = Mid(pCommand, 1, cSIZE_PANELID)
        frmMain.flxMES_Data.TextMatrix(23, 1) = .CF_PANELID
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        
        .CF_PANEL_OX_INFORMATION = Mid(pCommand, 1, cSIZE_CF_PNLINFO)
        frmMain.flxMES_Data.TextMatrix(24, 1) = .CF_PANEL_OX_INFORMATION
        pCommand = Mid(pCommand, cSIZE_CF_PNLINFO + 1)
        
        .PANEL_OWNER_TYPE = Mid(pCommand, 1, cSIZE_PANEL_OWNER_TYPE_MES)
        frmMain.flxMES_Data.TextMatrix(25, 1) = .PANEL_OWNER_TYPE
        pCommand = Mid(pCommand, cSIZE_PANEL_OWNER_TYPE_MES + 1)
        
        .ABNORMAL_CF = Mid(pCommand, 1, cSIZE_ABNORMAL_MES)
        frmMain.flxMES_Data.TextMatrix(26, 1) = .ABNORMAL_CF
        pCommand = Mid(pCommand, cSIZE_ABNORMAL_MES + 1)
        
        .ABNORMAL_TFT = Mid(pCommand, 1, cSIZE_ABNORMAL_MES)
        frmMain.flxMES_Data.TextMatrix(27, 1) = .ABNORMAL_TFT
        pCommand = Mid(pCommand, cSIZE_ABNORMAL_MES + 1)
        
        .ABNORMAL_LCD = Mid(pCommand, 1, cSIZE_ABNORMAL_LCD_MES)
        frmMain.flxMES_Data.TextMatrix(28, 1) = .ABNORMAL_LCD
        pCommand = Mid(pCommand, cSIZE_ABNORMAL_LCD_MES + 1)
        
        .GROUP_ID = Mid(pCommand, 1, cSIZE_GROUPID_MES)
        frmMain.flxMES_Data.TextMatrix(29, 1) = .GROUP_ID
        pCommand = Mid(pCommand, cSIZE_GROUPID_MES + 1)
        
'        .REPAIR_REWORK_COUNT = Mid(pCommand, 1, cSIZE_REWORKCOUNT_MES)
'        pCommand = Mid(pCommand, cSIZE_REWORKCOUNT_MES + 1)
        
        .POLARIZER_REWORK_COUNT = Mid(pCommand, 1, cSIZE_REWORKCOUNT_MES)
        frmMain.flxMES_Data.TextMatrix(34, 1) = .POLARIZER_REWORK_COUNT
        pCommand = Mid(pCommand, cSIZE_REWORKCOUNT_MES + 1)
        
        .LCD_Q_TAP_LOT_GROUPID = Mid(pCommand, 1, cSIZE_QTAPLOT_MES)
        frmMain.flxMES_Data.TextMatrix(39, 1) = .LCD_Q_TAP_LOT_GROUPID
        pCommand = Mid(pCommand, cSIZE_QTAPLOT_MES + 1)
        
        .SK_FLAG = Mid(pCommand, 1, cSIZE_SK_FLAG_MES)
        frmMain.flxMES_Data.TextMatrix(40, 1) = .SK_FLAG
        pCommand = Mid(pCommand, cSIZE_SK_FLAG_MES + 1)
        
        .CF_R_DEFECT_CODE = Mid(pCommand, 1, cSIZE_CF_R_DEFECT_CODE_MES)
        frmMain.flxMES_Data.TextMatrix(41, 1) = .CF_R_DEFECT_CODE
        pCommand = Mid(pCommand, cSIZE_CF_R_DEFECT_CODE_MES + 1)
        
        .ODK_AK_FLAG = Mid(pCommand, 1, cSIZE_ODK_AK_FLAG_MES)
        frmMain.flxMES_Data.TextMatrix(42, 1) = .ODK_AK_FLAG
        pCommand = Mid(pCommand, cSIZE_ODK_AK_FLAG_MES + 1)
        
        .BPAM_REWORK_FLAG = Mid(pCommand, 1, cSIZE_BPAM_REWORK_FLAG)
        frmMain.flxMES_Data.TextMatrix(43, 1) = .BPAM_REWORK_FLAG
        pCommand = Mid(pCommand, cSIZE_BPAM_REWORK_FLAG + 1)
        
        .LCD_BRIGHT_DOT_FLAG = Mid(pCommand, 1, cSIZE_LCD_BRIGHT_DOT_FLAG)
        frmMain.flxMES_Data.TextMatrix(44, 1) = .LCD_BRIGHT_DOT_FLAG
        pCommand = Mid(pCommand, cSIZE_LCD_BRIGHT_DOT_FLAG + 1)
        
        .CF_PS_HEIGHT_ERR_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(45, 1) = .CF_PS_HEIGHT_ERR_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .PI_INSPECTION_NG_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(46, 1) = .PI_INSPECTION_NG_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .PI_OVER_BAKE_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(47, 1) = .PI_OVER_BAKE_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .PI_OVER_Q_TIME_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(48, 1) = .PI_OVER_Q_TIME_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .ODF_OVER_BAKE_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(49, 1) = .ODF_OVER_BAKE_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .ODF_OVER_Q_TIME_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(50, 1) = .ODF_OVER_Q_TIME_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .HVA_OVER_BAKE_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(51, 1) = .HVA_OVER_BAKE_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .HVA_OVER_Q_TIME_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(52, 1) = .HVA_OVER_Q_TIME_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .SEAL_INSPECTION_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(53, 1) = .SEAL_INSPECTION_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .ODF_CHECKER_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(54, 1) = .ODF_CHECKER_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .ODF_DOOR_OPEN_FLAG = Mid(pCommand, 1, cSIZE_FLAG)
        frmMain.flxMES_Data.TextMatrix(55, 1) = .ODF_DOOR_OPEN_FLAG
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        
        .PRODUCTID = Mid(pCommand, 1, cSIZE_PRODUCTID_MES)
        frmMain.flxMES_Data.TextMatrix(58, 1) = .PRODUCTID
        pCommand = Mid(pCommand, cSIZE_PRODUCTID_MES + 1)
        
        .OWNERID = Mid(pCommand, 1, cSIZE_OWNERID_MES)
        frmMain.flxMES_Data.TextMatrix(59, 1) = .OWNERID
        pCommand = Mid(pCommand, cSIZE_OWNERID_MES + 1)
        
        .PREPROCESSID = Mid(pCommand, 1, cSIZE_PREPROCESSID)
        pCommand = Mid(pCommand, cSIZE_PREPROCESSID + 1)
        
        For intIndex = 1 To 9
            .SPARE(intIndex) = Mid(pCommand, 1, 25)
            frmMain.flxMES_Data.TextMatrix(59 + intIndex, 1) = .SPARE(intIndex)
            pCommand = Mid(pCommand, 26)
        Next intIndex
        
        .SPARE(10) = pCommand
        frmMain.flxMES_Data.TextMatrix(69, 1) = .SPARE(10)
        
        
      'For CATST PANEL ID SEND ERROR----Lucas
'        .PANELID = pubPANEL_INFO.AAAAAA
        
        'Read PFCD Data from Local Database
        Call Get_PFCD_DATA(typPFCD_PID, pPFCD)
        .X_TOTAL_PIXEL = typPFCD_PID.X_PIXEL_LENGTH
        frmMain.flxMES_Data.TextMatrix(35, 1) = .X_TOTAL_PIXEL
        .Y_TOTAL_PIXEL = typPFCD_PID.Y_PIXEL_LENGTH
        frmMain.flxMES_Data.TextMatrix(36, 1) = .Y_TOTAL_PIXEL
        .X_ONE_PIXEL_LENGTH = typPFCD_PID.DATA
        frmMain.flxMES_Data.TextMatrix(37, 1) = .X_ONE_PIXEL_LENGTH
        .Y_ONE_PIXEL_LENGTH = typPFCD_PID.GATE
        frmMain.flxMES_Data.TextMatrix(38, 1) = .Y_ONE_PIXEL_LENGTH
        
'Lucas 2011.12.26  Ver.0.7.33---Path change for CANRP/CALOI
        strRemote_Path = "LINK\" & "CATST\" & Mid(.PRODUCTID, 3, 5) & "\" & Left(.PANELID, 5) & "\" & Left(.PANELID, 8) & "\" & .PANELID & "\"
        strLocal_Path = App.PATH & "\GRADE_INFO\"
        strFileName = .PANELID & ".csv"
        Call Get_File_From_Host_by_Path(strRemote_Path, strLocal_Path, strFileName)
    End With
    
End Sub

Public Sub Decode_JOB_Information_Elements(ByVal pCommand As String, pJOB_DATA As JOB_DATA_STRUCTURE)

    With pJOB_DATA
        .CST_SEQUENCE = Left(pCommand, cSIZE_CST_SEQUENCE_JOB)
        pCommand = Mid(pCommand, cSIZE_CST_SEQUENCE_JOB + 1)
        
        .JOB_SEQUENCE = Left(pCommand, cSIZE_JOB_SEQUENCE_JOB)
        pCommand = Mid(pCommand, cSIZE_JOB_SEQUENCE_JOB + 1)
        
        .CIM_MODE = Left(pCommand, cSIZE_CIM_MODE_JOB)
        pCommand = Mid(pCommand, cSIZE_CIM_MODE_JOB + 1)
        
        .JOB_JUDGE = Left(pCommand, cSIZE_JOB_JUDGE_JOB)
        pCommand = Mid(pCommand, cSIZE_JOB_JUDGE_JOB + 1)
        
        .JOB_GRADE = Left(pCommand, cSIZE_JOB_GRADE_JOB)
        pCommand = Mid(pCommand, cSIZE_JOB_GRADE_JOB + 1)
        
        .GLASSID = Left(pCommand, cSIZE_GLASSID_JOB)
        pCommand = Mid(pCommand, cSIZE_GLASSID_JOB + 1)
        
        .BURR_CHECK_JUDGE = Left(pCommand, cSIZE_BURR_CHECK_JUDGE_JOB)
        pCommand = Mid(pCommand, cSIZE_BURR_CHECK_JUDGE_JOB + 1)
        
        .BEVELING_JUDGE = Left(pCommand, cSIZE_BEVELING_JUDGE_JOB)
        pCommand = Mid(pCommand, cSIZE_BEVELING_JUDGE_JOB + 1)
        
        .CLEANER_M_PORT_JUDGE = Left(pCommand, cSIZE_CLEANER_M_PORT_JUDGE_JOB)
        pCommand = Mid(pCommand, cSIZE_CLEANER_M_PORT_JUDGE_JOB + 1)
        
        .TEST_CV_JUDGE = Left(pCommand, cSIZE_TEST_CV_JUDGE_JOB)
        pCommand = Mid(pCommand, cSIZE_TEST_CV_JUDGE_JOB + 1)
        
        .SAMPLING_SLOT_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .PROCESS_INPUT_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .NEED_GRINDING_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .MISALIGNMENT_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .SMALL_MULTI_PANEL_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .AK_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .SK_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .NO_MATCH_GLASS_IN_BC_FLAG = Left(pCommand, cSIZE_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_FLAG_JOB + 1)
        
        .CASSETTE_SETTING_CODE = Left(pCommand, cSIZE_CASSETTE_SETTING_CODE_JOB)
        pCommand = Mid(pCommand, cSIZE_CASSETTE_SETTING_CODE_JOB + 1)
        
        .ABNORMAL_FLAG_CODE = Left(pCommand, cSIZE_ABNORMAL_FLAG_CODE_JOB)
        pCommand = Mid(pCommand, cSIZE_ABNORMAL_FLAG_CODE_JOB + 1)
        
        .LIGHT_ON_REASON_CODE = Left(pCommand, cSIZE_LIGHT_ON_REASON_CODE_JOB)
        pCommand = Mid(pCommand, cSIZE_LIGHT_ON_REASON_CODE_JOB + 1)
        
        .PANEL_NG_FLAG = Left(pCommand, cSIZE_PANEL_NG_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_PANEL_NG_FLAG_JOB + 1)
        
        .CUT_FLAG = Left(pCommand, cSIZE_CUT_FLAG_JOB)
        pCommand = Mid(pCommand, cSIZE_CUT_FLAG_JOB + 1)
        
        .RESERVED = Left(pCommand, cSIZE_RESERVED_JOB)
    End With
    
End Sub

Public Sub Decode_Share_Information_Elements(ByVal pCommand As String, pSHARE_DATA As SHARE_DATA_STRUCTURE)

    With pSHARE_DATA
        .PANELID = Left(pCommand, cSIZE_PANELID_SHARE)
        pCommand = Mid(pCommand, cSIZE_PANELID_SHARE + 1)
        
        .GLASS_TYPE = Left(pCommand, cSIZE_GLASS_TYPE_SHARE)
        pCommand = Mid(pCommand, cSIZE_GLASS_TYPE_SHARE + 1)
        
        .PRODUCTID = Left(pCommand, cSIZE_PRODUCTID_SHARE)
        pCommand = Mid(pCommand, cSIZE_PRODUCTID_SHARE + 1)
        
        .PROCESSID = Left(pCommand, cSIZE_PROCESSID_SHARE)
        pCommand = Mid(pCommand, cSIZE_PROCESSID_SHARE + 1)
        
        .RECIPEID = Left(pCommand, cSIZE_RECIPEID_SHARE)
        pCommand = Mid(pCommand, cSIZE_RECIPEID_SHARE + 1)
        
        .SALE_ORDER = Left(pCommand, cSIZE_SALE_ORDER_SHARE)
        pCommand = Mid(pCommand, cSIZE_SALE_ORDER_SHARE + 1)
        
        .CF_GLASSID = Left(pCommand, cSIZE_CF_GLASSID_SHARE)
        pCommand = Mid(pCommand, cSIZE_CF_GLASSID_SHARE + 1)
        
        .ARRAY_LOTID = Left(pCommand, cSIZE_ARRAY_LOTID_SHARE)
        pCommand = Mid(pCommand, cSIZE_ARRAY_LOTID_SHARE + 1)
        
        .ARRAY_GLASSID = Left(pCommand, cSIZE_ARRAY_GLASSID_SHARE)
        pCommand = Mid(pCommand, cSIZE_ARRAY_GLASSID_SHARE + 1)
        
        .CF_GLASS_INFO = Left(pCommand, cSIZE_CF_GLASS_INFO_SHARE)
        pCommand = Mid(pCommand, cSIZE_CF_GLASS_INFO_SHARE + 1)
        
        .TFT_PANEL_JUDGE = Left(pCommand, cSIZE_TFT_PANEL_JUDGE_SHARE)
        pCommand = Mid(pCommand, cSIZE_TFT_PANEL_JUDGE_SHARE + 1)
        
        .PRE_PROCESSID1 = Left(pCommand, cSIZE_PRE_PROCESSID1_SHARE)
        pCommand = Mid(pCommand, cSIZE_PRE_PROCESSID1_SHARE + 1)
        
        .GROUPID = Left(pCommand, cSIZE_GROUPID_SHARE)
        pCommand = Mid(pCommand, cSIZE_GROUPID_SHARE + 1)
        
        .TRANSFER_TIME = Left(pCommand, cSIZE_TRANSFER_TIME_SHARE)
    End With
    
End Sub

Public Sub Set_MES_Data(pCST_INFO As CST_INFO_ELEMENTS, pPNL_INFO As PANEL_INFO_ELEMENTS, pJOB_INFO As JOB_DATA_STRUCTURE, pSHARE_INFO As SHARE_DATA_STRUCTURE)

    With frmMain.flxMES_Data
'        .TextMatrix(0, 1) = pPNL_INFO.PANELID
'        .TextMatrix(1, 1) = pCST_INFO.PFCD
'        .TextMatrix(2, 1) = pSHARE_INFO.RECIPEID
'        .TextMatrix(3, 1) = pCST_INFO.PROCESS_NUM
'        .TextMatrix(4, 1) = pPNL_INFO.LIGHT_ON_PANEL_GRADE
'        .TextMatrix(5, 1) = pPNL_INFO.LIGHT_ON_REASON_CODE
'        .TextMatrix(6, 1) = pCST_INFO.OWNER
'        .TextMatrix(7, 1) = pPNL_INFO.CELL_LINE_RESCUE_FLAG
'        .TextMatrix(8, 1) = pPNL_INFO.SK_FLAG
'        .TextMatrix(9, 1) = pPNL_INFO.CF_PANELID
'        .TextMatrix(10, 1) = pPNL_INFO.CF_PANEL_OX_INFORMATION
'        .TextMatrix(11, 1) = pPNL_INFO.ABNORMAL_CF
'        .TextMatrix(12, 1) = pPNL_INFO.ABNORMAL_TFT
'        .TextMatrix(13, 1) = pPNL_INFO.ABNORMAL_LCD
'        .TextMatrix(14, 1) = pPNL_INFO.GROUP_ID
'        .TextMatrix(15, 1) = pPNL_INFO.SLOT_NUM
'        .TextMatrix(16, 1) = pPNL_INFO.LIGHT_ON_PANEL_GRADE
'        .TextMatrix(17, 1) = pPNL_INFO.LIGHT_ON_REASON_CODE
'        .TextMatrix(18, 1) = pPNL_INFO.CELL_LINE_RESCUE_FLAG
'        .TextMatrix(19, 1) = pPNL_INFO.CELL_REPAIR_JUDGE_GRADE
'        .TextMatrix(20, 1) = pPNL_INFO.TFT_REPAIR_GRADE
'        .TextMatrix(21, 1) = pPNL_INFO.REPAIR_REWORK_COUNT
'        .TextMatrix(22, 1) = pPNL_INFO.PRODUCTID
        .TextMatrix(20, 1) = pPNL_INFO.CELL_LINE_RESCUE_FLAG
        .TextMatrix(21, 1) = pPNL_INFO.CELL_REPAIR_JUDGE_GRADE
        .TextMatrix(30, 1) = pPNL_INFO.REPAIR_REWORK_COUNT
        .TextMatrix(31, 1) = pPNL_INFO.CARBONIZATION_FLAG
        .TextMatrix(32, 1) = pPNL_INFO.CARBONIZATION_GRADE
        .TextMatrix(33, 1) = pPNL_INFO.CARBONIZATION_REWORK_COUNT
    End With
    frmMain.lblPre_Judge.Caption = frmMain.flxMES_Data.TextMatrix(18, 1)
    frmMain.lblPre_Loss_Code.Caption = frmMain.flxMES_Data.TextMatrix(19, 1)
    frmMain.Repair.Caption = frmMain.flxMES_Data.TextMatrix(30, 1)
    frmJudge.Repair.Caption = frmMain.flxMES_Data.TextMatrix(30, 1)
End Sub

Public Sub Set_RUN_Data()

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strQuery                As String
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    
    Dim intRow                  As Integer
    
    Dim lngCurrent_Time         As Long
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        lngCurrent_Time = CLng(Format(TIME, "HHMMSS"))
        If (lngCurrent_Time <= 73000) And (lngCurrent_Time < 193000) Then
            strQuery = "SELECT * FROM PANEL_DATA WHERE "
            strQuery = strQuery & "RUN_DATE = " & CLng(Format(DATE, "YYYYMMDD")) & " AND "
            strQuery = strQuery & "RUN_TIME >= 73000 AND "
            strQuery = strQuery & "RUN_TIME < 193000"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveLast
                frmMain.flxRUN_Info.TextMatrix(2, 1) = lstRecord.RecordCount
            Else
                frmMain.flxRUN_Info.TextMatrix(2, 1) = "0"
            End If
            lstRecord.Close
        ElseIf lngCurrent_Time < 73000 Then
            strQuery = "SELECT * FROM PANEL_DATA WHERE "
            strQuery = strQuery & "RUN_DATE = " & CLng(Format(DATE - 1, "YYYYMMDD")) & " AND "
            strQuery = strQuery & "RUN_TIME >= 193000"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveLast
                frmMain.flxRUN_Info.TextMatrix(2, 1) = lstRecord.RecordCount
            Else
                frmMain.flxRUN_Info.TextMatrix(2, 1) = "0"
            End If
            lstRecord.Close
            
            strQuery = "SELECT * FROM PANEL_DATA WHERE "
            strQuery = strQuery & "RUN_DATE = " & CLng(Format(DATE, "YYYYMMDD"))
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveLast
                frmMain.flxRUN_Info.TextMatrix(2, 1) = CInt(frmMain.flxRUN_Info.TextMatrix(2, 1)) + lstRecord.RecordCount
            End If
            lstRecord.Close
        Else
            strQuery = "SELECT * FROM PANEL_DATA WHERE "
            strQuery = strQuery & "RUN_DATE = " & CLng(Format(DATE, "YYYYMMDD")) & " AND "
            strQuery = strQuery & "RUN_TIME >= 193000"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveLast
                frmMain.flxRUN_Info.TextMatrix(2, 1) = lstRecord.RecordCount
            Else
                frmMain.flxRUN_Info.TextMatrix(2, 1) = "0"
            End If
            lstRecord.Close
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Set_Average_Tact()

End Sub

Public Sub Send_Panel_Judge(ByVal pPanelID As String, ByVal pGrade As String, ByVal pLossCode As String, ByVal pSortFlag As String)

    Dim strCommand                  As String
    Dim strTemp                     As String
    Dim strState                    As String
    
    Dim intLength                   As Integer
    Dim intPortNo                   As Integer
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strState)
    
    If intPortNo > 0 Then
        strCommand = "QJPG"
        
        intLength = cSIZE_PANELID - Len(Trim(pPanelID))
        strCommand = strCommand & Trim(pPanelID) & Space(intLength)
        
        intLength = cSIZE_GRADE - Len(Trim(pGrade))
        strCommand = strCommand & Trim(pGrade) & Space(intLength)
        
        intLength = cSIZE_LOSSCODE - Len(Trim(pLossCode))
        strCommand = strCommand & Trim(pLossCode) & Space(intLength)
        
    '    intLength = cSIZE_FLAG - Len(Trim(pSortFlag))
    '    strCommand = strCommand & Trim(pSortFlag) & Space(intLength)
                
        If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
            If Len(EQP.Get_BackLight_Value) > 3 Then
                strCommand = strCommand & Left(EQP.Get_BackLight_Value, 3)
            Else
                intLength = cSIZE_BACKLIGHT_VALUE - Len(EQP.Get_BackLight_Value)
                strCommand = strCommand & Trim(EQP.Get_BackLight_Value) & Space(intLength)
            End If
        End If
        
        Call QUEUE.Put_Send_Command(intPortNo, strCommand)
    End If
    
End Sub

Public Function Get_RANK_DIVISION() As String

    Dim strPROD_TYPE                As String
    Dim strPROC_NUM                 As String
    Dim strPROD_ID                  As String
    
    strPROD_ID = frmMain.flxEQ_Information.TextMatrix(2, 1)
    If Len(strPROD_ID) > 5 Then
        Get_RANK_DIVISION = Left(strPROD_ID, 5)
    Else
        Get_RANK_DIVISION = strPROD_ID
    End If
    
    If frmMain.flxMES_Data.TextMatrix(2, 1) <> "" Then
        strPROD_TYPE = Left(frmMain.flxMES_Data.TextMatrix(2, 1), 1)
    Else
        strPROD_TYPE = ""
    End If
    
    strPROC_NUM = frmMain.flxMES_Data.TextMatrix(3, 1)
    Get_RANK_DIVISION = strPROD_TYPE & Get_RANK_DIVISION & strPROC_NUM
    
End Function

Public Function Check_Panel(pCST_INFO As CST_INFO_ELEMENTS, pPANEL_INFO As PANEL_INFO_ELEMENTS) As Integer

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    Dim strALTX                     As String
    
    Dim intFileNum                  As Integer
    Dim intFile_Index               As Integer
    Dim intIndex                    As Integer
    
    Dim bolFind                     As Boolean
    
    strPath = App.PATH & "\STANDARD_INFO\"
    intFile_Index = 0
    
    While (bolFind = False) And (intFile_Index <= 10)
        intFile_Index = intFile_Index + 1
        strFileName = "CheckPanelID" & intFile_Index & ".csv"
    Wend
    
End Function

Public Sub Get_Panel_Grade_LossCode(pGrade As String, pLossCode As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "KEYID = '" & RANK_OBJ.Get_Current_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With lstRecord
                pGrade = lstRecord.Fields("PANEL_GRADE")
                pLossCode = lstRecord.Fields("PANEL_LOSSCODE")
            End With
        End If
        
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Save_MES_Data(pCST_DATA As CST_INFO_ELEMENTS, pPANEL_DATA As PANEL_INFO_ELEMENTS, pJOB_DATA As JOB_DATA_STRUCTURE, pSHARE_DATA As SHARE_DATA_STRUCTURE)

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intIndex                    As Integer
    
    strPath = App.PATH & "\Env\"
    
    strFileName = "CST_DATA.mes"
    intFileNum = FreeFile
    Open strPath & strFileName For Output As intFileNum
    
    With pCST_DATA
        strTemp = "CSTID=" & .CSTID & vbCrLf
        strTemp = strTemp & "PFCD=" & .PFCD & vbCrLf
        strTemp = strTemp & "OWNER=" & .OWNER & vbCrLf
        strTemp = strTemp & "PROCESSNUM=" & .PROCESS_NUM & vbCrLf
        strTemp = strTemp & "PORTID=" & .PORTID & vbCrLf
        strTemp = strTemp & "PORTTYPE=" & .PORT_TYPE & vbCrLf
        strTemp = strTemp & "DESTINATION_FAB=" & .DESTINATION_FAB & vbCrLf
        strTemp = strTemp & "PANEL_COUNT=" & .PANEL_COUNT & vbCrLf
        strTemp = strTemp & "RMANO=" & .RMANO & vbCrLf
        strTemp = strTemp & "OQCNO=" & .OQCNO & vbCrLf
        strTemp = strTemp & "SOURCE_FAB=" & .SOURCE_FAB & vbCrLf
        For intIndex = 1 To 5
            strTemp = strTemp & "CST_SPARE" & intIndex & "=" & vbCrLf
        Next intIndex
        Print #intFileNum, strTemp
    End With
    Close intFileNum
    
    strFileName = "PANEL_DATA.mes"
    intFileNum = FreeFile
    Open strPath & strFileName For Output As intFileNum
    
    With pPANEL_DATA
        strTemp = "SLOT_NUM=" & .SLOT_NUM & vbCrLf
        strTemp = strTemp & "PANELID=" & .PANELID & vbCrLf
        strTemp = strTemp & "LIGHT_ON_PANEL_GRADE=" & .LIGHT_ON_PANEL_GRADE & vbCrLf
        strTemp = strTemp & "LIGHT_ON_REASON_CODE=" & .LIGHT_ON_REASON_CODE & vbCrLf
        strTemp = strTemp & "CELL_LINE_RESCUE_FLAG=" & .CELL_LINE_RESCUE_FLAG & vbCrLf
        strTemp = strTemp & "CELL_REPAIR_JUDGE_GRADE=" & .CELL_REPAIR_JUDGE_GRADE & vbCrLf
        strTemp = strTemp & "TFT_REPAIR_GRADE=" & .TFT_REPAIR_GRADE & vbCrLf
        strTemp = strTemp & "CF_PANELID=" & .CF_PANELID & vbCrLf
        strTemp = strTemp & "CF_PANEL_OX_INFORMATION=" & .CF_PANEL_OX_INFORMATION & vbCrLf
        strTemp = strTemp & "PANEL_OWNER_TYPE=" & .PANEL_OWNER_TYPE & vbCrLf
        strTemp = strTemp & "ABNORMAL_CF=" & .ABNORMAL_CF & vbCrLf
        strTemp = strTemp & "ABNORMAL_TFT=" & .ABNORMAL_TFT & vbCrLf
        strTemp = strTemp & "ABNORMAL_LCD=" & .ABNORMAL_LCD
        Print #intFileNum, strTemp
        
        strTemp = "GROUPID=" & .GROUP_ID & vbCrLf
        strTemp = strTemp & "REPAIR_REWORK_COUNT=" & .REPAIR_REWORK_COUNT & vbCrLf
        strTemp = strTemp & "CARBONIZATION_FLAG=" & .CARBONIZATION_FLAG & vbCrLf
        strTemp = strTemp & "CARBONIZATION_GRADE=" & .CARBONIZATION_GRADE & vbCrLf
        strTemp = strTemp & "CARBONIZATION_REWORK_COUNT=" & .CARBONIZATION_REWORK_COUNT & vbCrLf
        strTemp = strTemp & "POLARIZER_REWORK_COUNT=" & .POLARIZER_REWORK_COUNT & vbCrLf
        strTemp = strTemp & "X_TOTAL_PIXEL=" & .X_TOTAL_PIXEL & vbCrLf
        strTemp = strTemp & "Y_TOTAL_PIXEL=" & .Y_TOTAL_PIXEL & vbCrLf
        strTemp = strTemp & "X_ONE_PIXEL_LENGTH=" & .X_ONE_PIXEL_LENGTH & vbCrLf
        strTemp = strTemp & "Y_ONE_PIXEL_LENGTH=" & .Y_ONE_PIXEL_LENGTH & vbCrLf
        strTemp = strTemp & "LCD_Q_TAB_LOT_GROUPID=" & .LCD_Q_TAP_LOT_GROUPID & vbCrLf
        strTemp = strTemp & "SK_FLAG=" & .SK_FLAG & vbCrLf
        strTemp = strTemp & "CF_R_DEFECT_CODE=" & .CF_R_DEFECT_CODE & vbCrLf
        strTemp = strTemp & "ODF_AK_FLAG=" & .ODK_AK_FLAG & vbCrLf
        strTemp = strTemp & "BPAM_REWORK_FLAG=" & .BPAM_REWORK_FLAG & vbCrLf
        strTemp = strTemp & "LCD_BRIGHT_DOT_FLAG=" & .LCD_BRIGHT_DOT_FLAG
        Print #intFileNum, strTemp
        
        strTemp = "CF_PS_HEIGHT_ERR_FLAG=" & .CF_PS_HEIGHT_ERR_FLAG & vbCrLf
        strTemp = strTemp & "PI_INSPECTION_NG_FLAG=" & .PI_INSPECTION_NG_FLAG & vbCrLf
        strTemp = strTemp & "PI_OVER_BAKE_FLAG=" & .PI_OVER_BAKE_FLAG & vbCrLf
        strTemp = strTemp & "PI_OVER_Q_TIME_FLAG=" & .PI_OVER_Q_TIME_FLAG & vbCrLf
        strTemp = strTemp & "ODF_OVER_BAKE_FLAG=" & .ODF_OVER_BAKE_FLAG & vbCrLf
        strTemp = strTemp & "ODF_OVER_Q_TIME_FLAG=" & .ODF_OVER_Q_TIME_FLAG & vbCrLf
        strTemp = strTemp & "HVA_OVER_BAKE_FLAG=" & .HVA_OVER_BAKE_FLAG & vbCrLf
        strTemp = strTemp & "HVA_OVER_Q_TIME_FLAG=" & .HVA_OVER_Q_TIME_FLAG & vbCrLf
        strTemp = strTemp & "SEAL_INSPECTION_FLAG=" & .SEAL_INSPECTION_FLAG & vbCrLf
        strTemp = strTemp & "ODF_CHECK_FLAG=" & .ODF_CHECKER_FLAG & vbCrLf
        strTemp = strTemp & "ODF_DOOR_OPEN_FLAG=" & .ODF_DOOR_OPEN_FLAG & vbCrLf
        strTemp = strTemp & "LOT1_OPERATION_MODE=" & .LOT1_OPERATION_MODE & vbCrLf
        strTemp = strTemp & "LOT2_OPERATION_MODE=" & .LOT2_OPERATION_MODE & vbCrLf
        strTemp = strTemp & "PRODUCTID=" & .PRODUCTID & vbCrLf
        strTemp = strTemp & "OWNERID=" & .OWNERID & vbCrLf
        strTemp = strTemp & "PREPROCESSID=" & .PREPROCESSID & vbCrLf
        For intIndex = 1 To 9
            strTemp = strTemp & "SPARE" & intIndex & "=" & .SPARE(intIndex) & vbCrLf
        Next intIndex
        strTemp = strTemp & "SPARE10=" & .SPARE(10)
        Print #intFileNum, strTemp
    End With
    Close intFileNum
    
    strFileName = "JOB_DATA.mes"
    intFileNum = FreeFile
    Open strPath & strFileName For Output As intFileNum
    
    With pJOB_DATA
        strTemp = "CST_SEQUENCE=" & .CST_SEQUENCE & vbCrLf
        strTemp = strTemp & "JOB_SEQUENCE=" & .JOB_SEQUENCE & vbCrLf
        strTemp = strTemp & "CIM_MODE=" & .CIM_MODE & vbCrLf
        strTemp = strTemp & "JOB_JUDGE=" & .JOB_JUDGE & vbCrLf
        strTemp = strTemp & "JOB_GRADE=" & .JOB_GRADE & vbCrLf
        strTemp = strTemp & "GLASSID=" & .GLASSID & vbCrLf
        strTemp = strTemp & "BURR_CHECK_JUDGE=" & .BURR_CHECK_JUDGE & vbCrLf
        strTemp = strTemp & "BEVELING_JUDGE=" & .BEVELING_JUDGE & vbCrLf
        strTemp = strTemp & "CLEANER_M_PORT_JUDGE=" & .CLEANER_M_PORT_JUDGE & vbCrLf
        strTemp = strTemp & "TEST_CV_JUDGE=" & .TEST_CV_JUDGE & vbCrLf
        strTemp = strTemp & "SAMPLING_SLOT_FLAG=" & .SAMPLING_SLOT_FLAG & vbCrLf
        strTemp = strTemp & "PROCESS_INPUT_FLAG=" & .PROCESS_INPUT_FLAG
        Print #intFileNum, strTemp
        
        strTemp = "NEED_GRINDING_FLAG=" & .NEED_GRINDING_FLAG & vbCrLf
        strTemp = strTemp & "MISALIGNMENT_FLAG=" & .MISALIGNMENT_FLAG & vbCrLf
        strTemp = strTemp & "SMALL_MULTI_PANEL_FLAG=" & .SMALL_MULTI_PANEL_FLAG & vbCrLf
        strTemp = strTemp & "AK_FLAG=" & .AK_FLAG & vbCrLf
        strTemp = strTemp & "SK_FLAG=" & .SK_FLAG & vbCrLf
        strTemp = strTemp & "NO_MATCH_GLASS_IN_BC_FLAG=" & .NO_MATCH_GLASS_IN_BC_FLAG & vbCrLf
        strTemp = strTemp & "CASSETTING_SETTING_CODE=" & .CASSETTE_SETTING_CODE & vbCrLf
        strTemp = strTemp & "ABNORMAL_FLAG_CODE=" & .ABNORMAL_FLAG_CODE & vbCrLf
        strTemp = strTemp & "LIGHT_ON_REASON_CODE=" & .LIGHT_ON_REASON_CODE & vbCrLf
        strTemp = strTemp & "PANEL_NG_FLAG=" & .PANEL_NG_FLAG & vbCrLf
        strTemp = strTemp & "CUT_FLAG=" & .CUT_FLAG & vbCrLf
        strTemp = strTemp & "RESERVED=" & .RESERVED
        Print #intFileNum, strTemp
    End With
    Close intFileNum
    
    strFileName = "SHARE_DATA.mes"
    intFileNum = FreeFile
    Open strPath & strFileName For Output As intFileNum
    
    With pSHARE_DATA
        strTemp = "PANELID=" & .PANELID & vbCrLf
        strTemp = strTemp & "GLASS_TYPE=" & .GLASS_TYPE & vbCrLf
        strTemp = strTemp & "PRODUCTID=" & .PRODUCTID & vbCrLf
        strTemp = strTemp & "PROCESSID=" & .PROCESSID & vbCrLf
        strTemp = strTemp & "RECIPEID=" & .RECIPEID & vbCrLf
        strTemp = strTemp & "SALE_ORDER=" & .SALE_ORDER & vbCrLf
        strTemp = strTemp & "CF_GLASSID=" & .CF_GLASSID & vbCrLf
        strTemp = strTemp & "ARRAY_LOTID=" & .ARRAY_LOTID & vbCrLf
        strTemp = strTemp & "ARRAY_GLASSID=" & .ARRAY_GLASSID & vbCrLf
        strTemp = strTemp & "CF_GLASS_INFO=" & .CF_GLASS_INFO & vbCrLf
        strTemp = strTemp & "TFT_PANEL_JUDGE=" & .TFT_PANEL_JUDGE & vbCrLf
        strTemp = strTemp & "PRE_PRODESSID1=" & .PRE_PROCESSID1 & vbCrLf
        strTemp = strTemp & "GROUPID=" & .GROUPID & vbCrLf
        strTemp = strTemp & "TRANSFER_TIME=" & .TRANSFER_TIME
        Print #intFileNum, strTemp
    End With
    Close intFileNum
    
End Sub

Public Sub Get_MES_Data(pCST_DATA As CST_INFO_ELEMENTS, pPANEL_DATA As PANEL_INFO_ELEMENTS, pJOB_DATA As JOB_DATA_STRUCTURE, pSHARE_DATA As SHARE_DATA_STRUCTURE)

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    Dim strData                     As String
    
    Dim intFileNum                  As Integer
    Dim intIndex                    As Integer
    Dim intPosition                 As Integer
    
    strPath = App.PATH & "\Env\"
    
    strFileName = "CST_DATA.mes"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPosition = InStr(strTemp, "=")
            If intPosition > 0 Then
                With pCST_DATA
                    Select Case Left(strTemp, intPosition - 1)
                    Case "CSTID":
                        .CSTID = Mid(strTemp, intPosition + 1)
                    Case "PFCD":
                        .PFCD = Mid(strTemp, intPosition + 1)
                    Case "OWNER":
                        .OWNER = Mid(strTemp, intPosition + 1)
                    Case "PROCESSNUM"
                        .PROCESS_NUM = Mid(strTemp, intPosition + 1)
                    Case "PORTID":
                        .PORTID = Mid(strTemp, intPosition + 1)
                    Case "PORTTYPE":
                        .PORT_TYPE = Mid(strTemp, intPosition + 1)
                    Case "DESTINATION_FAB":
                        .DESTINATION_FAB = Mid(strTemp, intPosition + 1)
                    Case "PANEL_COUNT":
                        .PANEL_COUNT = Mid(strTemp, intPosition + 1)
                    Case "RMANO":
                        .RMANO = Mid(strTemp, intPosition + 1)
                    Case "OQCNO":
                        .OQCNO = Mid(strTemp, intPosition + 1)
                    Case "SOURCE_FAB":
                        .SOURCE_FAB = Mid(strTemp, intPosition + 1)
                    Case "CST_SPARE1":
                        .CST_SPARE(1) = Mid(strTemp, intPosition + 1)
                    Case "CST_SPARE2":
                        .CST_SPARE(2) = Mid(strTemp, intPosition + 1)
                    Case "CST_SPARE3":
                        .CST_SPARE(3) = Mid(strTemp, intPosition + 1)
                    Case "CST_SPARE4":
                        .CST_SPARE(4) = Mid(strTemp, intPosition + 1)
                    Case "CST_SPARE5":
                        .CST_SPARE(5) = Mid(strTemp, intPosition + 1)
                    End Select
                End With
            End If
        Wend
        Close intFileNum
    End If
    
    strFileName = "PANEL_DATA.mes"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPosition = InStr(strTemp, "=")
            If intPosition > 0 Then
                With pPANEL_DATA
                    Select Case Left(strTemp, intPosition - 1)
                    Case "SLOT_NUM":
                        .SLOT_NUM = Mid(strTemp, intPosition + 1)
                    Case "PANELID":
                        .PANELID = Mid(strTemp, intPosition + 1)
                    Case "LIGHT_ON_PANEL_GRADE":
                        .LIGHT_ON_PANEL_GRADE = Mid(strTemp, intPosition + 1)
                    Case "LIGHT_ON_REASON_CODE":
                        .LIGHT_ON_REASON_CODE = Mid(strTemp, intPosition + 1)
                    Case "CELL_LINE_RESCUE_FLAG":
                        .CELL_LINE_RESCUE_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CELL_REPAIR_JUDGE_GRADE":
                        .CELL_REPAIR_JUDGE_GRADE = Mid(strTemp, intPosition + 1)
                    Case "TFT_REPAIR_GRADE":
                        .TFT_REPAIR_GRADE = Mid(strTemp, intPosition + 1)
                    Case "CF_PANELID":
                        .CF_PANELID = Mid(strTemp, intPosition + 1)
                    Case "CF_PANEL_OX_INFORMATION":
                        .CF_PANEL_OX_INFORMATION = Mid(strTemp, intPosition + 1)
                    Case "PANEL_OWNER_TYPE":
                        .PANEL_OWNER_TYPE = Mid(strTemp, intPosition + 1)
                    Case "ABNORMAL_CF":
                        .ABNORMAL_CF = Mid(strTemp, intPosition + 1)
                    Case "ABNORMAL_TFT":
                        .ABNORMAL_TFT = Mid(strTemp, intPosition + 1)
                    Case "ABNORMAL_LCD":
                        .ABNORMAL_LCD = Mid(strTemp, intPosition + 1)
                    Case "GROUPID":
                        .GROUP_ID = Mid(strTemp, intPosition + 1)
                    Case "REPAIR_REWORK_COUNT":
                        .REPAIR_REWORK_COUNT = Mid(strTemp, intPosition + 1)
                    Case "CARBONIZATION_FLAG":
                        .CARBONIZATION_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CARBONIZATION_GRADE":
                        .CARBONIZATION_GRADE = Mid(strTemp, intPosition + 1)
                    Case "CARBONIZATION_REWORK_COUNT":
                        .CARBONIZATION_REWORK_COUNT = Mid(strTemp, intPosition + 1)
                    Case "POLARIZER_REWORK_COUNT":
                        .POLARIZER_REWORK_COUNT = Mid(strTemp, intPosition + 1)
                    Case "X_TOTAL_PIXEL":
                        .X_TOTAL_PIXEL = Mid(strTemp, intPosition + 1)
                    Case "Y_TOTAL_PIXEL":
                        .Y_TOTAL_PIXEL = Mid(strTemp, intPosition + 1)
                    Case "X_ONE_PIXEL_LENGTH":
                        .X_ONE_PIXEL_LENGTH = Mid(strTemp, intPosition + 1)
                    Case "Y_ONE_PIXEL_LENGTH":
                        .Y_ONE_PIXEL_LENGTH = Mid(strTemp, intPosition + 1)
                    Case "LCD_Q_TAB_LOT_GROUPID":
                        .LCD_Q_TAP_LOT_GROUPID = Mid(strTemp, intPosition + 1)
                    Case "SK_FLAG":
                        .SK_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CF_R_DEFECT_CODE":
                        .CF_R_DEFECT_CODE = Mid(strTemp, intPosition + 1)
                    Case "ODF_AK_FLAG":
                        .ODK_AK_FLAG = Mid(strTemp, intPosition + 1)
                    Case "BPAM_REWORK_FLAG":
                        .BPAM_REWORK_FLAG = Mid(strTemp, intPosition + 1)
                    Case "LCD_BRIGHT_DOT_FLAG":
                        .LCD_BRIGHT_DOT_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CF_PS_HEIGHT_ERR_FLAG":
                        .CF_PS_HEIGHT_ERR_FLAG = Mid(strTemp, intPosition + 1)
                    Case "PI_INSPECTION_NG_FLAG":
                        .PI_INSPECTION_NG_FLAG = Mid(strTemp, intPosition + 1)
                    Case "PI_OVER_BAKE_FLAG":
                        .PI_OVER_BAKE_FLAG = Mid(strTemp, intPosition + 1)
                    Case "PI_OVER_Q_TIME_FLAG":
                        .PI_OVER_Q_TIME_FLAG = Mid(strTemp, intPosition + 1)
                    Case "ODF_OVER_BAKE_FLAG":
                        .ODF_OVER_BAKE_FLAG = Mid(strTemp, intPosition + 1)
                    Case "ODF_OVER_Q_TIME_FLAG":
                        .ODF_OVER_Q_TIME_FLAG = Mid(strTemp, intPosition + 1)
                    Case "HVA_OVER_BAKE_FLAG":
                        .HVA_OVER_BAKE_FLAG = Mid(strTemp, intPosition + 1)
                    Case "HVA_OVER_Q_TIME_FLAG":
                        .HVA_OVER_Q_TIME_FLAG = Mid(strTemp, intPosition + 1)
                    Case "SEAL_INSPECTION_FLAG":
                        .SEAL_INSPECTION_FLAG = Mid(strTemp, intPosition + 1)
                    Case "ODF_CHECK_FLAG":
                        .ODF_CHECKER_FLAG = Mid(strTemp, intPosition + 1)
                    Case "ODF_DOOR_OPEN_FLAG":
                        .ODF_DOOR_OPEN_FLAG = Mid(strTemp, intPosition + 1)
                    Case "LOT1_OPERATION_MODE":
                        .LOT1_OPERATION_MODE = Mid(strTemp, intPosition + 1)
                    Case "LOT2_OPERATION_MODE":
                        .LOT2_OPERATION_MODE = Mid(strTemp, intPosition + 1)
                    Case "PRODUCTID":
                        .PRODUCTID = Mid(strTemp, intPosition + 1)
                    Case "OWNERID":
                        .OWNERID = Mid(strTemp, intPosition + 1)
                    Case "PREPROCESSID":
                        .PREPROCESSID = Mid(strTemp, intPosition + 1)
                    Case "SPARE1":
                        .SPARE(1) = Mid(strTemp, intPosition + 1)
                    Case "SPARE2":
                        .SPARE(2) = Mid(strTemp, intPosition + 1)
                    Case "SPARE3":
                        .SPARE(3) = Mid(strTemp, intPosition + 1)
                    Case "SPARE4":
                        .SPARE(4) = Mid(strTemp, intPosition + 1)
                    Case "SPARE5":
                        .SPARE(5) = Mid(strTemp, intPosition + 1)
                    Case "SPARE6":
                        .SPARE(6) = Mid(strTemp, intPosition + 1)
                    Case "SPARE7":
                        .SPARE(7) = Mid(strTemp, intPosition + 1)
                    Case "SPARE8":
                        .SPARE(8) = Mid(strTemp, intPosition + 1)
                    Case "SPARE9":
                        .SPARE(9) = Mid(strTemp, intPosition + 1)
                    Case "SPARE10":
                        .SPARE(10) = Mid(strTemp, intPosition + 1)
                    End Select
                End With
            End If
        Wend
        Close intFileNum
    End If
    
    strFileName = "JOB_DATA.mes"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPosition = InStr(strTemp, "=")
            If intPosition > 0 Then
                With pJOB_DATA
                    Select Case Left(strTemp, intPosition - 1)
                    Case "CST_SEQUENCE":
                        .CST_SEQUENCE = Mid(strTemp, intPosition + 1)
                    Case "JOB_SEQUENCE":
                        .JOB_SEQUENCE = Mid(strTemp, intPosition + 1)
                    Case "CIM_MODE":
                        .CIM_MODE = Mid(strTemp, intPosition + 1)
                    Case "JOB_JUDGE":
                        .JOB_JUDGE = Mid(strTemp, intPosition + 1)
                    Case "JOB_GRADE":
                        .JOB_GRADE = Mid(strTemp, intPosition + 1)
                    Case "GLASSID":
                        .GLASSID = Mid(strTemp, intPosition + 1)
                    Case "BURR_CHECK_JUDGE":
                        .BURR_CHECK_JUDGE = Mid(strTemp, intPosition + 1)
                    Case "BEVELING_JUDGE":
                        .BEVELING_JUDGE = Mid(strTemp, intPosition + 1)
                    Case "CLEANER_M_PORT_JUDGE":
                        .CLEANER_M_PORT_JUDGE = Mid(strTemp, intPosition + 1)
                    Case "TEST_CV_JUDGE":
                        .TEST_CV_JUDGE = Mid(strTemp, intPosition + 1)
                    Case "SAMPLING_SLOT_FLAG":
                        .SAMPLING_SLOT_FLAG = Mid(strTemp, intPosition + 1)
                    Case "PROCESS_INPUT_FLAG":
                        .PROCESS_INPUT_FLAG = Mid(strTemp, intPosition + 1)
                    Case "NEED_GRINDING_FLAG":
                        .NEED_GRINDING_FLAG = Mid(strTemp, intPosition + 1)
                    Case "MISALIGNMENT_FLAG":
                        .MISALIGNMENT_FLAG = Mid(strTemp, intPosition + 1)
                    Case "SMALL_MULTI_PANEL_FLAG":
                        .SMALL_MULTI_PANEL_FLAG = Mid(strTemp, intPosition + 1)
                    Case "AK_FLAG":
                        .AK_FLAG = Mid(strTemp, intPosition + 1)
                    Case "SK_FLAG":
                        .SK_FLAG = Mid(strTemp, intPosition + 1)
                    Case "NO_MATCH_GLASS_IN_BC_FLAG":
                        .NO_MATCH_GLASS_IN_BC_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CASSETTING_SETTING_CODE":
                        .CASSETTE_SETTING_CODE = Mid(strTemp, intPosition + 1)
                    Case "ABNORMAL_FLAG_CODE":
                        .ABNORMAL_FLAG_CODE = Mid(strTemp, intPosition + 1)
                    Case "LIGHT_ON_REASON_CODE":
                        .LIGHT_ON_REASON_CODE = Mid(strTemp, intPosition + 1)
                    Case "PANEL_NG_FLAG":
                        .PANEL_NG_FLAG = Mid(strTemp, intPosition + 1)
                    Case "CUT_FLAG":
                        .CUT_FLAG = Mid(strTemp, intPosition + 1)
                    Case "RESERVED":
                        .RESERVED = Mid(strTemp, intPosition + 1)
                    End Select
                End With
            End If
        Wend
        Close intFileNum
    End If
    
    strFileName = "SHARE_DATA.mes"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPosition = InStr(strTemp, "=")
            If intPosition > 0 Then
                strData = Mid(strTemp, intPosition + 1)
                With pSHARE_DATA
                    Select Case Left(strTemp, intPosition - 1)
                    Case "PANELID":
                        .PANELID = strData
                    Case "GLASS_TYPE":
                        .GLASS_TYPE = strData
                    Case "PRODUCTID":
                        .PRODUCTID = strData
                    Case "PROCESSID":
                        .PROCESSID = strData
                    Case "RECIPEID":
                        .RECIPEID = strData
                    Case "SALE_ORDER":
                        .SALE_ORDER = strData
                    Case "CF_GLASSID":
                        .CF_GLASSID = strData
                    Case "ARRAY_LOTID":
                        .ARRAY_LOTID = strData
                    Case "ARRAY_GLASSID":
                        .ARRAY_GLASSID = strData
                    Case "CF_GLASS_INFO":
                        .CF_GLASS_INFO = strData
                    Case "TFT_PANEL_JUDGE":
                        .TFT_PANEL_JUDGE = strData
                    Case "PRE_PRODESSID1":
                        .PRE_PROCESSID1 = strData
                    Case "GROUPID":
                        .GROUPID = strData
                    Case "TRANSFER_TIME":
                        .TRANSFER_TIME = strData
                    End Select
                End With
            End If
        Wend
        Close intFileNum
    End If
    
End Sub

Public Sub Insert_Panel_MES_Data(pPANEL_DATA As PANEL_DATA, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                 pJOB_MES_DATA As JOB_DATA_STRUCTURE, pSHARE_MES_DATA As SHARE_DATA_STRUCTURE)
                                 
    Dim dbMyDB                      As Database
    
    Dim strQuery                    As String
    
    Dim intIndex                    As Integer
    
    If Dir(pPANEL_DATA.PATH & pPANEL_DATA.FILENAME, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(pPANEL_DATA.PATH & pPANEL_DATA.FILENAME)
        
        With pCST_MES_DATA
            strQuery = "INSERT INTO CST_MES_DATA VALUES ("
            strQuery = strQuery & "'" & .CSTID & "', "
            strQuery = strQuery & "'" & .PFCD & "', "
            strQuery = strQuery & "'" & .OWNER & "', "
            strQuery = strQuery & "'" & .PROCESS_NUM & "', "
            strQuery = strQuery & "'" & .PORTID & "', "
            strQuery = strQuery & "'" & .PORT_TYPE & "', "
            strQuery = strQuery & "'" & .DESTINATION_FAB & "', "
            strQuery = strQuery & "'" & .PANEL_COUNT & "', "
            strQuery = strQuery & "'" & .RMANO & "', "
            strQuery = strQuery & "'" & .OQCNO & "', "
            strQuery = strQuery & "'" & .SOURCE_FAB & "', "
            For intIndex = 1 To 4
                strQuery = strQuery & "'" & .CST_SPARE(intIndex) & "', "
            Next intIndex
            strQuery = strQuery & "'" & .CST_SPARE(5) & "')"
            
            dbMyDB.Execute strQuery
        End With
        
        With pPANEL_MES_DATA
            strQuery = "INSERT INTO PANEL_MES_DATA VALUES ("
            strQuery = strQuery & "'" & .SLOT_NUM & "', "
            strQuery = strQuery & "'" & .PANELID & "', "
            strQuery = strQuery & "'" & .LIGHT_ON_PANEL_GRADE & "', "
            strQuery = strQuery & "'" & .LIGHT_ON_REASON_CODE & "', "
            strQuery = strQuery & "'" & .CELL_LINE_RESCUE_FLAG & "', "
            strQuery = strQuery & "'" & .CELL_REPAIR_JUDGE_GRADE & "', "
            strQuery = strQuery & "'" & .TFT_REPAIR_GRADE & "', "
            strQuery = strQuery & "'" & .CF_PANELID & "', "
            strQuery = strQuery & "'" & .CF_PANEL_OX_INFORMATION & "', "
            strQuery = strQuery & "'" & .PANEL_OWNER_TYPE & "', "
            strQuery = strQuery & "'" & .ABNORMAL_CF & "', "
            strQuery = strQuery & "'" & .ABNORMAL_TFT & "', "
            strQuery = strQuery & "'" & .ABNORMAL_LCD & "', "
            strQuery = strQuery & "'" & .GROUP_ID & "', "
            strQuery = strQuery & "'" & .REPAIR_REWORK_COUNT & "', "
            strQuery = strQuery & "'" & .CARBONIZATION_FLAG & "', "
            strQuery = strQuery & "'" & .CARBONIZATION_GRADE & "', "
            strQuery = strQuery & "'" & .CARBONIZATION_REWORK_COUNT & "', "
            strQuery = strQuery & "'" & .POLARIZER_REWORK_COUNT & "', "
            strQuery = strQuery & "'" & .X_TOTAL_PIXEL & "', "
            strQuery = strQuery & "'" & .Y_TOTAL_PIXEL & "', "
            strQuery = strQuery & "'" & .X_ONE_PIXEL_LENGTH & "', "
            strQuery = strQuery & "'" & .Y_ONE_PIXEL_LENGTH & "', "
            strQuery = strQuery & "'" & .LCD_Q_TAP_LOT_GROUPID & "', "
            strQuery = strQuery & "'" & .SK_FLAG & "', "
            strQuery = strQuery & "'" & .CF_R_DEFECT_CODE & "', "
            strQuery = strQuery & "'" & .ODK_AK_FLAG & "', "
            strQuery = strQuery & "'" & .BPAM_REWORK_FLAG & "', "
            strQuery = strQuery & "'" & .LCD_BRIGHT_DOT_FLAG & "', "
            strQuery = strQuery & "'" & .CF_PS_HEIGHT_ERR_FLAG & "', "
            strQuery = strQuery & "'" & .PI_INSPECTION_NG_FLAG & "', "
            strQuery = strQuery & "'" & .PI_OVER_BAKE_FLAG & "', "
            strQuery = strQuery & "'" & .PI_OVER_Q_TIME_FLAG & "', "
            strQuery = strQuery & "'" & .ODF_OVER_BAKE_FLAG & "', "
            strQuery = strQuery & "'" & .ODF_OVER_Q_TIME_FLAG & "', "
            strQuery = strQuery & "'" & .HVA_OVER_BAKE_FLAG & "', "
            strQuery = strQuery & "'" & .HVA_OVER_Q_TIME_FLAG & "', "
            strQuery = strQuery & "'" & .SEAL_INSPECTION_FLAG & "', "
            strQuery = strQuery & "'" & .ODF_CHECKER_FLAG & "', "
            strQuery = strQuery & "'" & .ODF_DOOR_OPEN_FLAG & "', "
            strQuery = strQuery & "'" & .LOT1_OPERATION_MODE & "', "
            strQuery = strQuery & "'" & .LOT2_OPERATION_MODE & "', "
            strQuery = strQuery & "'" & .PRODUCTID & "', "
            strQuery = strQuery & "'" & .OWNERID & "', "
            strQuery = strQuery & "'" & .PREPROCESSID & "', "
            For intIndex = 1 To 9
                strQuery = strQuery & "'" & .SPARE(intIndex) & "', "
            Next intIndex
            strQuery = strQuery & "'" & .SPARE(10) & "')"
            
            dbMyDB.Execute strQuery
        End With
        
        With pJOB_MES_DATA
            strQuery = "INSERT INTO JOB_DATA VALUES ("
            strQuery = strQuery & "'" & .CST_SEQUENCE & "', "
            strQuery = strQuery & "'" & .JOB_SEQUENCE & "', "
            strQuery = strQuery & "'" & .CIM_MODE & "', "
            strQuery = strQuery & "'" & .JOB_JUDGE & "', "
            strQuery = strQuery & "'" & .JOB_GRADE & "', "
            strQuery = strQuery & "'" & .GLASSID & "', "
            strQuery = strQuery & "'" & .BURR_CHECK_JUDGE & "', "
            strQuery = strQuery & "'" & .BEVELING_JUDGE & "', "
            strQuery = strQuery & "'" & .CLEANER_M_PORT_JUDGE & "', "
            strQuery = strQuery & "'" & .TEST_CV_JUDGE & "', "
            strQuery = strQuery & "'" & .SAMPLING_SLOT_FLAG & "', "
            strQuery = strQuery & "'" & .PROCESS_INPUT_FLAG & "', "
            strQuery = strQuery & "'" & .NEED_GRINDING_FLAG & "', "
            strQuery = strQuery & "'" & .MISALIGNMENT_FLAG & "', "
            strQuery = strQuery & "'" & .SMALL_MULTI_PANEL_FLAG & "', "
            strQuery = strQuery & "'" & .AK_FLAG & "', "
            strQuery = strQuery & "'" & .SK_FLAG & "', "
            strQuery = strQuery & "'" & .NO_MATCH_GLASS_IN_BC_FLAG & "', "
            strQuery = strQuery & "'" & .CASSETTE_SETTING_CODE & "', "
            strQuery = strQuery & "'" & .ABNORMAL_FLAG_CODE & "', "
            strQuery = strQuery & "'" & .LIGHT_ON_REASON_CODE & "', "
            strQuery = strQuery & "'" & .PANEL_NG_FLAG & "', "
            strQuery = strQuery & "'" & .CUT_FLAG & "', "
            strQuery = strQuery & "'" & .RESERVED & "')"
            
            dbMyDB.Execute strQuery
        End With
        
        With pSHARE_MES_DATA
            strQuery = "INSERT INTO SHARED_DATA VALUES ("
            strQuery = strQuery & "'" & .PANELID & "', "
            strQuery = strQuery & "'" & .GLASS_TYPE & "', "
            strQuery = strQuery & "'" & .PRODUCTID & "', "
            strQuery = strQuery & "'" & .PROCESSID & "', "
            strQuery = strQuery & "'" & .RECIPEID & "', "
            strQuery = strQuery & "'" & .SALE_ORDER & "', "
            strQuery = strQuery & "'" & .CF_GLASSID & "', "
            strQuery = strQuery & "'" & .ARRAY_LOTID & "', "
            strQuery = strQuery & "'" & .ARRAY_GLASSID & "', "
            strQuery = strQuery & "'" & .CF_GLASS_INFO & "', "
            strQuery = strQuery & "'" & .TFT_PANEL_JUDGE & "', "
            strQuery = strQuery & "'" & .PRE_PROCESSID1 & "', "
            strQuery = strQuery & "'" & .GROUPID & "', "
            strQuery = strQuery & "'" & .TRANSFER_TIME & "')"
            
            dbMyDB.Execute strQuery
        End With
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Read_TFT_CF_PanelID()

    Dim dbMyDB                  As Database
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strAlarm_Msg            As String
    
    Dim intFileIndex            As Integer
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    
    Dim bolFirst_Line           As Boolean
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        For intFileIndex = 1 To 10
            strPath = App.PATH & "\Env\STANDARD_INFO\"
            strFileName = "checkTFTPanelID" & intFileIndex & ".csv"
            
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                intFileNum = FreeFile
                Open strPath & strFileName For Input As intFileNum
                
                bolFirst_Line = True
                While Not EOF(intFileNum)
                    Line Input #intFileNum, strTemp
                    If bolFirst_Line = True Then
                        intPos = InStr(strTemp, ",")
                        If intPos > 0 Then
                            strAlarm_Msg = Left(strTemp, intPos - 1)
                        Else
                            strAlarm_Msg = strTemp
                        End If
                        bolFirst_Line = False
                    Else
                        strQuery = "INSERT INTO ABNORMAL_PANEL VALUES ("
                        strQuery = strQuery & "'" & strAlarm_Msg & "', "
                        strQuery = strQuery & "'" & Left(strTemp, cSIZE_PANELID) & "', "
                        strQuery = strQuery & "'" & Left(strTemp, 1) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                Wend
                
                Close intFileNum
            End If
        Next intFileIndex
        
        For intFileIndex = 1 To 10
            strPath = App.PATH & "\Env\STANDARD_INFO\"
            strFileName = "checkCFPanelID" & intFileIndex & ".csv"
            
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                intFileNum = FreeFile
                Open strPath & strFileName For Input As intFileNum
                
                bolFirst_Line = True
                While Not EOF(intFileNum)
                    Line Input #intFileNum, strTemp
                    If bolFirst_Line = True Then
                        intPos = InStr(strTemp, ",")
                        If intPos > 0 Then
                            strAlarm_Msg = Left(strTemp, intPos - 1)
                        Else
                            strAlarm_Msg = strTemp
                        End If
                        bolFirst_Line = False
                    Else
                        strQuery = "INSERT INTO ABNORMAL_PANEL VALUES ("
                        strQuery = strQuery & "'" & strAlarm_Msg & "', "
                        strQuery = strQuery & "'" & Left(strTemp, cSIZE_PANELID) & "', "
                        strQuery = strQuery & "'" & Left(strTemp, 1) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                Wend
                
                Close intFileNum
            End If
        Next intFileIndex
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_TFT_CF_PanelID", ErrMsg)
    
End Sub

Public Sub Read_Check_MES_DATA()

    Dim dbMyDB                  As Database
    
    Dim arrDATA(1 To 53)          As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strAlarm_Msg            As String
    
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    Dim intLoop_Count           As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "CheckMesData.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If UCase(Left(strTemp, 5)) <> "ALARM" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDATA(intIndex) = Left(strTemp, intPos - 1)
                            If arrDATA(intIndex) = "" Then
                                arrDATA(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDATA(intIndex) = strTemp
                        intLoop_Count = intIndex
                        
                        strQuery = "INSERT INTO ABNORMAL_MES_DATA VALUES ("
                        For intIndex = 1 To (intLoop_Count - 1)
                            strQuery = strQuery & "'" & arrDATA(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDATA(intIndex) & "')"
                    
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_Check_MES_DATA", ErrMsg)
    
End Sub

Public Sub Read_Assign_Grade()

    Dim dbMyDB                  As Database
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strPFCD                 As String
    Dim strPROCESSNUM           As String
    Dim strDESTINATION_FAB      As String
    Dim strNew_Grade            As String
    Dim strPanelID              As String
        
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intPriority             As String
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        For intIndex = 1 To 5
            strPath = App.PATH & "\Env\STANDARD_INFO\"
            strFileName = "AssignGrade" & intIndex & ".csv"
            
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                intFileNum = FreeFile
                Open strPath & strFileName For Input As intFileNum
                
                While Not EOF(intFileNum)
                    Line Input #intFileNum, strTemp
                    intPos = InStr(strTemp, ",")
                    If intPos > 0 Then
                        If Left(strTemp, 4) <> "PFCD" Then
 'Lucas Ver.0.9.19 2012.04.01=============================For assign Grade mdb Changed
                          If Len(Left(strTemp, intPos - 1)) <> cSIZE_PANELID Then
 'Lucas Ver.0.9.19 2012.04.01=============================For assign Grade mdb Changed
                            strPFCD = Left(strTemp, intPos - 1)
                            If strPFCD = "" Then
                                strPFCD = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            strPROCESSNUM = Left(strTemp, intPos - 1)
                            If strPROCESSNUM = "" Then
                                strPROCESSNUM = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            strDESTINATION_FAB = Left(strTemp, intPos - 1)
                            If strDESTINATION_FAB = "" Then
                                strDESTINATION_FAB = " "
                            End If
                            strNew_Grade = Mid(strTemp, intPos + 1)
                            If strNew_Grade = "" Then
                                strNew_Grade = " "
                            End If
                            intPriority = intIndex
                         Else
 'Lucas Ver.0.9.19 2012.04.01=============================For assign Grade mdb Changed
                            strPanelID = Left(strTemp, cSIZE_PANELID)
 'Lucas Ver.0.9.19 2012.04.01=============================For assign Grade mdb Changed
                            strQuery = "INSERT INTO ASSIGN_GRADE VALUES ("
                            strQuery = strQuery & "'" & strPFCD & "', "
                            strQuery = strQuery & "'" & strPanelID & "', "
                            strQuery = strQuery & "'" & strPROCESSNUM & "', "
                            strQuery = strQuery & "'" & strDESTINATION_FAB & "', "
                            strQuery = strQuery & "'" & strNew_Grade & "', "
                            strQuery = strQuery & intPriority & ")"
                            
                            dbMyDB.Execute (strQuery)
                        End If
                        End If
                    End If
                Wend
                Close intFileNum
            Else
                Call SaveLog("Read_Assign_Grade", strPath & strFileName & " is not found.")
            End If
        Next intIndex
        
        dbMyDB.Close
    Else
        Call MsgBox("STANDARD_INFO.mdb file is not exist.", vbOKOnly, "DB File not found.")
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_Assign_Grade", ErrMsg)
    
End Sub

Public Sub Read_Control()

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strFlagPath             As String
    Dim strFlagFileName         As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strITEM_NAME            As String
    Dim strITEM_VALUE           As String
    
    Dim intFileNum              As Integer
    Dim intFlagFileNum          As Integer
    Dim intPos                  As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

'*******************************************
'Date :2012.02.06
'issued by K.H.KIM
'-Delete standard_info.mdb file delete command
'*******************************************

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
'    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
'        Kill strDB_Path & strDB_FileName
'    End If
'    FileCopy strDB_Path & "STANDARD_INFO_temp.mdb", strDB_Path & strDB_FileName
    Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
    strPath = App.PATH & "\Env\STANDARD_INFO\"
    strFileName = "Control.csv"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
                        
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                If UCase(Left(strTemp, intPos - 1)) <> "ITEMS" Then
                    If UCase(Left(strTemp, intPos - 1)) = "TIMEFLAG" Then
                        strFlagPath = App.PATH & "\Env\"
                        strFlagFileName = "DownloadFlag.dat"
                        intFlagFileNum = FreeFile
                        
                        strTemp = Mid(strTemp, intPos + 1)
                        intPos = InStr(strTemp, ",")
                        Open strFlagPath & strFlagFileName For Output As intFlagFileNum
                        
                        If intPos > 0 Then
                            Print #intFlagFileNum, Left(strTemp, intPos - 1)
                            Call ENV.Set_Download_Flag(Left(strTemp, intPos - 1))
                        Else
                            Print #intFlagFileNum, strTemp
                            Call ENV.Set_Download_Flag(strTemp)
                        End If
                        
                        Close intFlagFileNum
                    Else
                        strITEM_NAME = Left(strTemp, intPos - 1)
                        strITEM_VALUE = Mid(strTemp, intPos + 1, 1)
                        
                        strQuery = "SELECT * FROM ITEM_CONTROL WHERE "
                        strQuery = strQuery & "ITEM_NAME = '" & strITEM_NAME & "'"
                        
                        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                        
                        If lstRecord.EOF = True Then
                            lstRecord.Close
                            
                            strQuery = "INSERT INTO ITEM_CONTROL VALUES ("
                            strQuery = strQuery & "'" & strITEM_NAME & "', "
                            strQuery = strQuery & "'" & strITEM_VALUE & "')"
                            
                            dbMyDB.Execute (strQuery)
                        Else
                            lstRecord.Close
                            
                            strQuery = "UPDATE ITEM_CONTROL SET "
                            strQuery = strQuery & "USES = '" & strITEM_VALUE & "' WHERE "
                            strQuery = strQuery & "ITEM_NAME = '" & strITEM_NAME & "'"
                            
                            dbMyDB.Execute (strQuery)
                        End If
                    End If
                End If
            End If
        Wend
        
        Close intFileNum
    End If
    
    dbMyDB.Close
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_Control", ErrMsg)
    
End Sub

Public Sub Read_PreJudgeGradeChange1()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 10) As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strPFCD                 As String
    Dim strPROCESSNUM           As String
    Dim strNew_Grade            As String
    Dim strData                 As String
    Dim strGate                 As String
    Dim strPre_Grade            As String
    
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intLimit_Left           As Integer
    Dim intLimit_Right          As Integer
    Dim intLimit_Upper          As Integer
    Dim intLimit_Bottom         As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PreJudgeGradeChange1.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        strPFCD = Left(strTemp, intPos - 1)
                        If strPFCD = "" Then
                            strPFCD = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strData = Left(strTemp, intPos - 1)
                        If strData = "" Then
                            strData = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strGate = Left(strTemp, intPos - 1)
                        If strGate = "" Then
                            strGate = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strPROCESSNUM = Left(strTemp, intPos - 1)
                        If strPROCESSNUM = "" Then
                            strPROCESSNUM = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strPre_Grade = Left(strTemp, intPos - 1)
                        If strPre_Grade = "" Then
                            strPre_Grade = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strNew_Grade = Left(strTemp, intPos - 1)
                        If strNew_Grade = "" Then
                            strNew_Grade = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intLimit_Upper = 0
                        Else
                            intLimit_Upper = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intLimit_Bottom = 0
                        Else
                            intLimit_Bottom = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intLimit_Left = 0
                        Else
                            intLimit_Left = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intLimit_Right = 0
                        Else
                            intLimit_Right = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO PRE_JUDGE_CHANGE_GRADE1 VALUES ("
                        strQuery = strQuery & "'" & strPFCD & "', "
                        strQuery = strQuery & "'" & strData & "', "
                        strQuery = strQuery & "'" & strGate & "', "
                        strQuery = strQuery & "'" & strPROCESSNUM & "', "
                        strQuery = strQuery & "'" & strPre_Grade & "', "
                        strQuery = strQuery & "'" & strNew_Grade & "', "
                        strQuery = strQuery & intLimit_Upper & ", "
                        strQuery = strQuery & intLimit_Bottom & ", "
                        strQuery = strQuery & intLimit_Left & ", "
                        strQuery = strQuery & intLimit_Right & ", "
                        For intIndex = 1 To 9
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(10) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PreJudgeGradeChange1", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PreJudgeGradeChange1", ErrMsg)
    
End Sub

Public Sub Read_PreJudgeGradeChange2()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 3)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strPFCD                 As String
    Dim strPROCESSNUM           As String
    Dim strNew_Grade            As String
    Dim strData                 As String
    Dim strGate                 As String
    Dim strPre_Grade            As String
    
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intLimit_Count          As Integer
    Dim intTotal_Division       As Integer
    Dim intDevide_Division      As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PreJudgeGradeChange2.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        strPFCD = Left(strTemp, intPos - 1)
                        If strPFCD = "" Then
                            strPFCD = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strData = Left(strTemp, intPos - 1)
                        If strData = "" Then
                            strData = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strGate = Left(strTemp, intPos - 1)
                        If strGate = "" Then
                            strGate = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strPROCESSNUM = Left(strTemp, intPos - 1)
                        If strPROCESSNUM = "" Then
                            strPROCESSNUM = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strPre_Grade = Left(strTemp, intPos - 1)
                        If strPre_Grade = "" Then
                            strPre_Grade = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        strNew_Grade = Left(strTemp, intPos - 1)
                        If strNew_Grade = "" Then
                            strNew_Grade = " "
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intLimit_Count = 0
                        Else
                            intLimit_Count = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intTotal_Division = 0
                        Else
                            intTotal_Division = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If intPos = 1 Then
                            intDevide_Division = 0
                        Else
                            intDevide_Division = CInt(Left(strTemp, intPos - 1))
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO PRE_JUDGE_CHANGE_GRADE2 VALUES ("
                        strQuery = strQuery & "'" & strPFCD & "', "
                        strQuery = strQuery & "'" & strData & "', "
                        strQuery = strQuery & "'" & strGate & "', "
                        strQuery = strQuery & "'" & strPROCESSNUM & "', "
                        strQuery = strQuery & "'" & strPre_Grade & "', "
                        strQuery = strQuery & "'" & strNew_Grade & "', "
                        strQuery = strQuery & intLimit_Count & ", "
                        strQuery = strQuery & intTotal_Division & ", "
                        strQuery = strQuery & intDevide_Division & ", "
                        strQuery = strQuery & "'" & arrDEFECT_CODE(1) & "', "
                        strQuery = strQuery & "'" & arrDEFECT_CODE(2) & "', "
                        strQuery = strQuery & "'" & arrDEFECT_CODE(3) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PreJudgeGradeChange2", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PreJudgeGradeChange2", ErrMsg)
    
End Sub

Public Sub Read_PreJudgeGradeChange3()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 6)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PreJudgeGradeChange3.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO PRE_JUDGE_CHANGE_GRADE3 VALUES ("
                        For intIndex = 1 To 5
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(6) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PreJudgeGradeChange3", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PreJudgeGradeChange3", ErrMsg)

End Sub

Public Sub Read_PostJudgeOtherRule1()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 6)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PostJudgeOtherRule1.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO POST_JUDGE_OTHER_RULE1 VALUES ("
                        For intIndex = 1 To 5
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PostJudgeOtherRule1", strPath & strFileName & " is not found.")
        End If
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PostJudgeOtherRule1", ErrMsg)

End Sub

Public Sub Read_PostJudgeOtherRule2()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 5)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PostJudgeOtherRule2.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO POST_JUDGE_OTHER_RULE2 VALUES ("
                        For intIndex = 1 To 4
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PostJudgeOtherRule2", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PostJudgeOtherRule2", ErrMsg)

End Sub

Public Sub Read_PostJudgeOtherRule3()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 5)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PostJudgeOtherRule3.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO POST_JUDGE_OTHER_RULE3 VALUES ("
                        For intIndex = 1 To 4
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PostJudgeOtherRule3", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PostJudgeOtherRule3", ErrMsg)

End Sub

Public Sub Read_PostJudgeGradeChange1()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 6)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PostJudgeGradeChange1.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO POST_JUDGE_GRADE_CHANGE1 VALUES ("
                        For intIndex = 1 To 5
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PostJudgeGradeChange1", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PostJudgeGradeChange1", ErrMsg)

End Sub

Public Sub Read_PostJudgeGradeChange2()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 6)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "PostJudgeGradeChange2.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO POST_JUDGE_GRADE_CHANGE2 VALUES ("
                        For intIndex = 1 To 5
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_PostJudgeGradeChange2", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_PostJudgeGradeChange2", ErrMsg)

End Sub

Public Sub Read_CheckPanelIDChangeGrade()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 15)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strPanelID              As String
    
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intFileIndex            As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        For intFileIndex = 1 To 5
            strFileName = "CheckPanelIDChangeGrade" & intFileIndex & ".csv"
            
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                intFileNum = FreeFile
                Open strPath & strFileName For Input As intFileNum
                
                While Not EOF(intFileNum)
                    Line Input #intFileNum, strTemp
                    intPos = InStr(strTemp, ",")
                    If intPos > 0 Then
                        If Left(strTemp, 4) <> "PFCD" Then
                            If Len(Left(strTemp, intPos - 1)) <> cSIZE_PANELID Then
                                intIndex = 0
                                While intPos > 0
                                    intIndex = intIndex + 1
                                    arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                                    If arrDEFECT_CODE(intIndex) = "" Then
                                        arrDEFECT_CODE(intIndex) = " "
                                    End If
                                    strTemp = Mid(strTemp, intPos + 1)
                                    intPos = InStr(strTemp, ",")
                                Wend
                                intIndex = intIndex + 1
                                If strTemp = "" Then
                                    strTemp = " "
                                End If
                                arrDEFECT_CODE(intIndex) = strTemp
                            Else
                            'Lucas 2012.03.20 Ver.0.9.16-------------Add no change Grade Num
                                arrDEFECT_CODE(14) = Left(strTemp, cSIZE_PANELID)
                                arrDEFECT_CODE(15) = "CheckPanelIDChangeGrade" & intFileIndex
                                strQuery = "INSERT INTO CHECK_PANELID_CHANGE_GRADE VALUES ("
                                'Lucas 2012.03.20 Ver.0.9.16-------------Add no change Grade Num
                                For intIndex = 1 To 14
                                    strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                                Next intIndex
                                strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                                
                                dbMyDB.Execute (strQuery)
                            End If
                        End If
'                    Else
'                        If strTemp = "" Then
'                            strTemp = " "
'                        End If
'                        arrDEFECT_CODE(9) = strTemp
'
'                        strQuery = "INSERT INTO CHECK_PANELID_CHANGE_GRADE VALUES ("
'                        For intIndex = 1 To 8
'                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
'                        Next intIndex
'                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
'
'                        dbMyDB.Execute (strQuery)
                    End If
                Wend
                Close intFileNum
            Else
                Call SaveLog("Read_CheckPanelIDChangeGrade", strPath & strFileName & " is not found.")
            End If
        Next intFileIndex
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_CheckPanelIDChangeGrade", ErrMsg)

End Sub

Public Sub Read_ChangeGrade()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 11) As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "ChangeGrade.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
'==========================================================================================================
'
'  Modify Date : 2011. 12. 20
'  Modify by K.H. KIM
'  Content
'    - If index number is less than array size check remained data
'
'
'  Start of modify
'
'==========================================================================================================
                        If intIndex <= 11 Then
                            If strTemp = "" Then
                                strTemp = " "
                            End If
                            arrDEFECT_CODE(intIndex) = strTemp
                        End If
'===========================================================================================================
'
'  End of modify
'
'===========================================================================================================
                        strQuery = "INSERT INTO CHANGE_GRADE VALUES ("
                        For intIndex = 1 To 10
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_ChangeGrade", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_ChangeGrade", ErrMsg)

End Sub

Public Sub Read_ChangeGradeByDefectCode()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 6)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "ChangeGradeByDefectCode.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While (intPos > 0) And (intIndex < 5)
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        intPos = InStr(strTemp, ",")
                        If intPos > 0 Then
                            strTemp = Left(strTemp, intPos - 1)
                        End If
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO CHANGE_GRADE_DEFECT_CODE VALUES ("
                        For intIndex = 1 To 5
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_ChangeGradeByDefectCode", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_ChangeGradeByDefectCode", ErrMsg)

End Sub

Public Sub Read_RepairPointTimes()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 9)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "RepairPointTimes.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO REPAIR_POINT_TIMES VALUES ("
                        For intIndex = 1 To 8
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_RepairPointTimes", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_RepairPointTimes", ErrMsg)

End Sub

Public Sub Read_FlagChangeGrade()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 56) As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "FlagChangeGrade.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If Left(strTemp, 4) <> "PFCD" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                            If arrDEFECT_CODE(intIndex) = "" Then
                                arrDEFECT_CODE(intIndex) = " "
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO FLAG_CHANGE_GRADE VALUES ("
                        For intIndex = 1 To 55
                            strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_FlagChangeGrade", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_FlagChangeGrade", ErrMsg)

End Sub

Public Sub Read_SKChange()

    Dim dbMyDB                  As Database
    
    Dim arrDEFECT_CODE(1 To 7)  As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String

    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intIndex                As Integer
    Dim intSampling_Value       As Integer
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strPath = App.PATH & "\Env\STANDARD_INFO\"
        strFileName = "SKchange.csv"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    If UCase(Left(strTemp, 7)) <> "MACHINE" Then
                        intIndex = 0
                        While intPos > 0
                            intIndex = intIndex + 1
                            If intIndex <> 5 Then
                                arrDEFECT_CODE(intIndex) = Left(strTemp, intPos - 1)
                                If arrDEFECT_CODE(intIndex) = "" Then
                                    arrDEFECT_CODE(intIndex) = " "
                                End If
                            Else
                                If intPos = 1 Then
                                    intSampling_Value = 0
                                Else
                                    intSampling_Value = CInt(Left(strTemp, intPos - 1))
                                End If
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intIndex = intIndex + 1
                        If strTemp = "" Then
                            strTemp = " "
                        End If
                        arrDEFECT_CODE(intIndex) = strTemp
                        
                        strQuery = "INSERT INTO SK_CHANGE VALUES ("
                        For intIndex = 1 To 6
                            If intIndex <> 5 Then
                                strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "', "
                            Else
                                strQuery = strQuery & intSampling_Value & ", "
                            End If
                        Next intIndex
                        strQuery = strQuery & "'" & arrDEFECT_CODE(intIndex) & "')"
                        
                        dbMyDB.Execute (strQuery)
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call SaveLog("Read_SKChange", strPath & strFileName & " is not found.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Read_SKChange", ErrMsg)

End Sub

Public Sub Read_Rank_Data(ByVal pDB_FileName As String)
    Dim typCST_INFO                 As CST_INFO_ELEMENTS
    Dim typRANK_DATA                As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA(1 To 50)      As GRADE_DATA_STRUCTURE
    
    Dim arrGrade(1 To 50)           As String
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    Dim intCol                      As Integer
    Dim intIndex                    As Integer
    Dim intArray_Index              As Integer
    Dim intArray_Count              As Integer
    Dim intFile_Count               As Integer
    '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

    
   
    If (pubCST_INFO.PFCD <> EQP.Get_Pre_PFCD) Or (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Or (typCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Then
        strPath = App.PATH & "\Env\Standard_Info\"
        Call Clear_DEFECT_LIST
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = pDB_FileName
        
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Kill strDB_Path & strDB_FileName
        End If
        FileCopy strDB_Path & "RANK_temp.mdb", strDB_Path & strDB_FileName
        Call SaveLog("Read_Rank_Data", strDB_Path & strDB_FileName & " create.")
        
       'Lucas 2012.01.05 ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
       If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
            strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
            Call Get_File_From_Host(strFileName, "RankSet")
       Else:
           strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
           Call Get_File_From_Host(strFileName, "RankSet")
       End If
                '==========================================END
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            Select Case Mid(strFileName, 7, 4)
            Case "3000":
                typRANK_DATA.RANK_DIVISION = "LOI-1"
            Case "3650":
                typRANK_DATA.RANK_DIVISION = "RLOI-1"
            Case "4600":
                typRANK_DATA.RANK_DIVISION = "LOI-2"
            Case "4650":
                typRANK_DATA.RANK_DIVISION = "RLOI-2"
            End Select
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    If InStr(Left(strTemp, intPos - 1), "CODE") = 0 Then
                    
                        With typRANK_DATA
                            .DEFECT_CODE = Left(strTemp, intPos - 1)
                            .DEFECT_DIVISION = Mid(.DEFECT_CODE, 2, 1)
                            
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DEFECT_NAME = Left(strTemp, intPos - 1)
                            If Len(.DEFECT_NAME) > 8 Then
                                .DEFECT_NAME = Left(.DEFECT_NAME, 8)
                            End If
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DEFECT_TYPE = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .JUDGE_OR_NOT = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .USE_XY = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DETAIL_DIVISION = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .ACCUMULATION = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .ADDRESS_COUNT = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                           '============Leo 2012.05.22 Add Rank Level Start
                            For intRankLevel = 0 To UBound(RankLevel)
                                intPos = InStr(strTemp, ",")
                                .Rank(intRankLevel) = Left(strTemp, intPos - 1)
                                strTemp = Mid(strTemp, intPos + 1)
                            Next intRankLevel
                            
'                            intPos = InStr(strTemp, ",")
'                            .RANK_Y = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
'
'                            intPos = InStr(strTemp, ",")
'                            .RANK_L = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
'
'                            intPos = InStr(strTemp, ",")
'                            .RANK_K = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
'
'                            intPos = InStr(strTemp, ",")
'                            .RANK_C = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
'
'                            intPos = InStr(strTemp, ",")
'                            .RANK_S = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
                            '============Leo 2012.05.22 Add Rank Level End
'                            intPos = InStr(strTemp, ",")
'                            .ODF = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .PRIORITY = CInt(Left(strTemp, intPos - 1))
                            strTemp = Mid(strTemp, intPos + 1)
                            
'                            intPos = InStr(strTemp, ",")
'                            .POP_UP = Left(strTemp, intPos - 1)
'                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                        End With
                        intArray_Index = 0
                        While intPos > 0
                            intArray_Index = intArray_Index + 1
                            With typGRADE_DATA(intArray_Index)
                                .DEFECT_CODE = typRANK_DATA.DEFECT_CODE
                                .GRADE = arrGrade(intArray_Index)
                                .RANK = Left(strTemp, intPos - 1)
                                .RANK_DIVISION = typRANK_DATA.RANK_DIVISION
                            End With
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intArray_Index = intArray_Index + 1
                        With typGRADE_DATA(intArray_Index)
                            .DEFECT_CODE = typRANK_DATA.DEFECT_CODE
                            .GRADE = arrGrade(intArray_Index)
                            .RANK = strTemp
                            .RANK_DIVISION = typRANK_DATA.RANK_DIVISION
                        End With
                        Call Insert_Rank_Data(typRANK_DATA, typGRADE_DATA, intArray_Index)
                        Call Insert_DEFECT_LIST(typRANK_DATA)
                    Else
                        strTemp = Mid(strTemp, intPos + 1)
                        intCol = 1
                        
                        intPos = InStr(strTemp, ",")
                        While UCase(Left(strTemp, intPos - 1)) <> "PRIORITY"
                            strTemp = Mid(strTemp, intPos + 1)
                            intCol = intCol + 1
                            intPos = InStr(strTemp, ",")
                        Wend
                        strTemp = Mid(strTemp, intPos + 1)
                        intCol = intCol + 1
                        
                        intPos = InStr(strTemp, ",")
                        intArray_Index = 0
                        While intPos > 0
                            intArray_Index = intArray_Index + 1
                            arrGrade(intArray_Index) = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            intPos = InStr(strTemp, ",")
                        Wend
                        intArray_Index = intArray_Index + 1
                        intArray_Count = intArray_Index
                        arrGrade(intArray_Index) = strTemp
                        RANK_OBJ.Set_Highest_Grade (arrGrade(1))
                    End If
                End If
            Wend
            Close intFileNum
        Else
            Call Show_Message("Rank file not found", strFileName & " does not exist.")
        End If
        
        Call RANK_OBJ.Init_RANK_Priority
'        Call ENV.Reset_Download_Flag
    End If
    
End Sub

Private Sub Insert_Rank_Data(pRANK_DATA As RANK_DATA_STRUCTURE, pGRADE_DATA() As GRADE_DATA_STRUCTURE, ByVal pGRADE_COUNT As Integer)

    Dim dbMyDB                      As Database
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intIndex                    As Integer
    
       '============Leo 2012.05.22 Add Rank Level Start
    Dim intCount                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

    
    strDB_Path = App.PATH & "\DB\"
               'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
    If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
            strDB_FileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    Else
            strDB_FileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    End If
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        
        With pRANK_DATA
            strQuery = "INSERT INTO RANK_DATA (RANK_DIVISION,DEFECT_CODE,DEFECT_NAME,DEFECT_DIVISION,DEFECT_TYPE,JUDGE_OR_NOT,USE_XY,DETAIL_DIVISION,ACCUMULATION,ADDRESS_COUNT,ODF,PRIORITY,POP_UP "
    '============Leo 2012.05.22 Add Rank Level Start
            For intCount = 0 To UBound(RankLevel)
                strQuery = strQuery & " , RANK_" & RankLevel(intCount)
            Next intCount
     '============Leo 2012.05.22 Add Rank Level end
            strQuery = strQuery & " ) VALUES ("
            strQuery = strQuery & "'" & .RANK_DIVISION & "', "
            strQuery = strQuery & "'" & .DEFECT_CODE & "', "
            strQuery = strQuery & "'" & .DEFECT_NAME & "', "
            strQuery = strQuery & "'" & .DEFECT_DIVISION & "', "
            strQuery = strQuery & "'" & .DEFECT_TYPE & "', "
            strQuery = strQuery & "'" & .JUDGE_OR_NOT & "', "
            strQuery = strQuery & "'" & .USE_XY & "', "
            strQuery = strQuery & "'" & .DETAIL_DIVISION & "', "
            strQuery = strQuery & "'" & .ACCUMULATION & "', "
            strQuery = strQuery & "'" & .ADDRESS_COUNT & "', "
            strQuery = strQuery & "'" & .ODF & "', "
            strQuery = strQuery & .PRIORITY & ", "
            strQuery = strQuery & "'" & .POP_UP & "'"
       '============Leo 2012.05.22 Add Rank Level Start
            For intCount = 0 To UBound(.Rank)
                If .Rank(intCount) <> "" Then
                    strQuery = strQuery & ",'" & .Rank(intCount) & "' "
                End If
            Next intCount
         '============Leo 2012.05.22 Add Rank Level End
            
            strQuery = strQuery & ")"
        End With
        dbMyDB.Execute (strQuery)
        
        For intIndex = 1 To pGRADE_COUNT
            With pGRADE_DATA(intIndex)
                strQuery = "INSERT INTO GRADE_DATA VALUES ("
                strQuery = strQuery & "'" & .RANK_DIVISION & "', "
                strQuery = strQuery & "'" & .DEFECT_CODE & "', "
                strQuery = strQuery & "'" & .GRADE & "', "
                strQuery = strQuery & "'" & .RANK & "')"
            End With
            dbMyDB.Execute (strQuery)
        Next intIndex
        
        dbMyDB.Close
    End If
    
End Sub

Private Sub Clear_DEFECT_LIST()

    Dim dbMyDB                  As Database
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "DELETE FROM DEFECT_LIST"
        
        dbMyDB.Execute (strQuery)
        
        dbMyDB.Close
        
        DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Parameter_Temp.mdb", dbLangChineseSimplified
        Kill strDB_Path & strDB_FileName
        Name strDB_Path & "Parameter_Temp.mdb" As strDB_Path & strDB_FileName
        
    End If
    
End Sub

Private Sub Insert_DEFECT_LIST(pRANK_DATA As RANK_DATA_STRUCTURE)

    Dim dbMyDB                  As Database
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        With pRANK_DATA
            strQuery = "INSERT INTO DEFECT_LIST VALUES ("
            strQuery = strQuery & "'" & .DEFECT_CODE & "', "
            strQuery = strQuery & "'" & .DEFECT_NAME & "', "
            strQuery = strQuery & "'" & Mid(.DEFECT_CODE, 2, 1) & "')"
        End With
        
        dbMyDB.Execute (strQuery)
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Reset_Interlock()

    With frmJudge
        .flxDefect_A.Enabled = True
        .flxDefect_B.Enabled = True
        .flxDefect_C.Enabled = True
        .flxDefect_D.Enabled = True
        .flxDefect_E.Enabled = True
        .flxDefect_F.Enabled = True
        .flxDefect_G.Enabled = True
        .flxDefect_H.Enabled = True
        .flxDefect_I.Enabled = True
    End With
    
End Sub

Public Sub DELETE_STANDARD_INFO(ByVal pTABLE_NAME As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "SELECT * FROM " & pTABLE_NAME
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            
            strQuery = "DELETE FROM " & pTABLE_NAME
            
            dbMyDB.Execute (strQuery)
        Else
            lstRecord.Close
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Public Function Make_CST_DATA_DB(ByVal pMonth As Integer, ByVal pDay As Integer, pPANEL_DATA As PANEL_DATA, pPANEL_INFO As PANEL_INFO_ELEMENTS) As Boolean

    Dim dbMyDB                      As Database
  
    Dim strDB_Path                  As String
    Dim strRemote_Path              As String
    Dim strFileName                 As String
    Dim strSource_Path              As String
    Dim strData_FileName            As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim FTP_OBJ                     As New clsFTP
    Dim Sourcepath                  As String
    Dim SourceName                  As String
    Dim intFileNum                  As Integer
    Dim strTemp                     As String
    
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    If Dir(strDB_Path, vbDirectory) = "" Then
        MkDir strDB_Path
    End If
    strDB_Path = strDB_Path & pMonth & "\"
    If Dir(strDB_Path, vbDirectory) = "" Then
        MkDir strDB_Path
    End If
    strDB_Path = strDB_Path & pDay & "\"
    If Dir(strDB_Path, vbDirectory) = "" Then
        MkDir strDB_Path
    End If
    
    strSource_Path = App.PATH & "\DB\"
    strData_FileName = pPANEL_DATA.PANELID & ".mdb"
    
    If Dir(strDB_Path & strData_FileName, vbNormal) <> "" Then
        Kill strDB_Path & strData_FileName
    End If
    
    FileCopy strSource_Path & "Panel_Data_Temp.mdb", strDB_Path & strData_FileName
    
    strDB_FileName = "Result.mdb"
    
    If Dir(strSource_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strSource_Path & strDB_FileName)
        
        With pPANEL_DATA
            strQuery = "INSERT INTO PANEL_DATA VALUES ("
            strQuery = strQuery & "'" & .KEYID & "', "
            strQuery = strQuery & "'" & .TIME & "', "
            strQuery = strQuery & "'" & .PANELID & "', "
            strQuery = strQuery & "'" & .PANEL_RANK & "', "
            strQuery = strQuery & "'" & .PANEL_GRADE & "', "
            strQuery = strQuery & "'" & .PANEL_LOSSCODE & "', "
            strQuery = strQuery & "'" & .LOSSCODE_NAME & "', "
            strQuery = strQuery & "'" & .USER_NAME & "', "
            strQuery = strQuery & "'" & .PANEL_TYPE1 & "', "
            strQuery = strQuery & "'" & .PANEL_TYPE2 & "', "
            strQuery = strQuery & "'" & strDB_Path & "', "
            strQuery = strQuery & "'" & strData_FileName & "', "
            strQuery = strQuery & .RUN_DATE & ", "
            strQuery = strQuery & .RUN_TIME & ", "
            strQuery = strQuery & .TACT_TIME & ")"
            .PATH = strDB_Path
            .FILENAME = strData_FileName
        End With
        dbMyDB.Execute (strQuery)
        
        dbMyDB.Close
    End If
    Call RANK_OBJ.Set_Current_KEYID(pPANEL_DATA.KEYID)
    
'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - When a current equipment is CATST/Operator mode skip the function share file download
'
'
'  Start of modify
'
'==========================================================================================================

'    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) <> "CATST" Then      'The equipment is CALOI
        strRemote_Path = "Link\" & "CATST\" & Mid(pPANEL_INFO.PRODUCTID, 3, 5) & "\" & Mid(pPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pPANEL_INFO.PANELID, 1, 8) & "\" & pPANEL_INFO.PANELID & "\"
        strFileName = pPANEL_INFO.PANELID & ".csv"
        Call Get_File_From_Host_by_Path(strRemote_Path, strDB_Path, strFileName)
        Call Read_Share_Defect_File(pPANEL_INFO, strDB_Path, strFileName)
'==========================================================================================================
'
'  Modify Date : 2012. 03. 20
'  Modify by Lucas
'  Content
'    - Get the address from Source data and show them on the Frmjudge
'    - Delete the Operator/CATST.Because it didn't need to download the share data.
'    - Add Source download and show it to the FrmJudge
'  Start of modify
'
'==========================================================================================================
        strRemote_Path = "CATST\" & Left(pPANEL_INFO.PRODUCTID, 11) & "0" & "\" & Mid(pPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pPANEL_INFO.PANELID, 1, 8) & "\" & pPANEL_INFO.PANELID & "\" & "BACKUP\"
    If FTP_OBJ.Init_FTP_Client = True Then
            Call FTP_OBJ.Open_Session
            strFileName = FTP_OBJ.FTP_Get_FileList("*.CSV", strRemote_Path)
          If strFileName <> "" Then
            Sourcepath = App.PATH & "\ENV\"
            If Dir(Sourcepath & strFileName, vbNormal) <> "" Then
                intFileNum = FreeFile
                Open Sourcepath & strFileName For Input As intFileNum
                While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                SourceName = strTemp
                Wend
                Close #intFileNum
            End If
          End If
           FTP_OBJ.Close_Session
           FTP_OBJ.Disconnect_FTP_Client
    End If
        
If SourceName <> "" Then
                 Call Get_File_From_Host_by_Path(strRemote_Path, strDB_Path, SourceName)
                 Call Read_Source_Defect_File(pPANEL_INFO, strDB_Path, SourceName)
End If

frmJudge.Polarizor.Caption = pPANEL_INFO.POLARIZER_REWORK_COUNT
        
'    Else
'        If frmMain.flxEQ_Information.TextMatrix(3, 1) = "Operator" Then         'The equipment is CATST and operation mode is operator mode
'            strRemote_Path = "Link\" & "CATST\" & Mid(pPANEL_INFO.PRODUCTID, 3, 5) & "\" & Mid(pPANEL_INFO.PANELID, 1, 5) & "\"
'            strRemote_Path = strRemote_Path & Mid(pPANEL_INFO.PANELID, 1, 8) & "\" & pPANEL_INFO.PANELID & "\"
'            strFileName = pPANEL_INFO.PANELID & ".csv"
'                Call Get_File_From_Host_by_Path(strRemote_Path, strDB_Path, strFileName)
'            Call Read_Share_Defect_File(pPANEL_INFO, strDB_Path, strFileName)
'        End If
'   End If
    
'==========================================================================================================
'
'  End of modify
'
'==========================================================================================================
    Make_CST_DATA_DB = True
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Make_CST_DATA_DB", ErrMsg)
    
    Make_CST_DATA_DB = False
    
End Function

Public Sub Read_Share_Defect_File(pPANEL_INFO As PANEL_INFO_ELEMENTS, ByVal pLocalPath As String, ByVal pFileName As String)

    Dim strTemp                             As String
    Dim strTitle                            As String
    
    Dim intFileNum                          As Integer
    Dim intPos                              As Integer
    Dim intRework_Count_Pos                 As Integer
    Dim intCell_Line_Rescue_Flag_Pos        As Integer
    Dim intCell_Repair_Judge_Grade_Pos      As Integer
    Dim intCarbonization_Flag_Pos           As Integer
    Dim intCarbonization_Grade_Pos          As Integer
    Dim intCarbonization_Rework_Count_Pos   As Integer
    Dim intItem_Index                       As Integer
    Dim intIndex                            As Integer
    
    Dim bolRead_Data                        As Boolean

    If Dir(pLocalPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open pLocalPath & pFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            
            If InStr(UCase(strTemp), "CANRP_PANEL_DEFECT_DATA_BEGIN") > 0 Then
                bolRead_Data = False
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    intItem_Index = 0
                    While (intPos > 0) And (bolRead_Data = False)
                        intItem_Index = intItem_Index + 1
                        strTitle = Left(strTemp, intPos - 1)
                        Select Case UCase(strTitle)
                        Case "CELL_REPAIR_GRADE":
                            intCell_Repair_Judge_Grade_Pos = intItem_Index
                        Case "LINE_RESCUE_FLAG":
                            intCell_Line_Rescue_Flag_Pos = intItem_Index
                        Case "LASERREPAIR_COUNT":
                            intRework_Count_Pos = intItem_Index
                        End Select
                        
                        strTemp = Mid(strTemp, intPos + 1)
                        intPos = InStr(strTemp, ",")
                    Wend
                    intItem_Index = intItem_Index + 1
                    Select Case UCase(strTemp)
                    Case "CELL_REPAIR_GRADE":
                        intCell_Repair_Judge_Grade_Pos = intItem_Index
                    Case "LINE_RESCUE_FLAG":
                        intCell_Line_Rescue_Flag_Pos = intItem_Index
                    Case "LASERREPAIR_COUNT":
                        intRework_Count_Pos = intItem_Index
                    End Select
                    Line Input #intFileNum, strTemp
                    While UCase(strTemp) <> "CANRP_PANEL_DEFECT_DATA_END"
                        intPos = InStr(strTemp, ",")
                        If intPos > 0 Then
                            intIndex = 0
                            While intPos > 0
                                intIndex = intIndex + 1
                                Select Case intIndex
                                Case intCell_Repair_Judge_Grade_Pos:
                                    pPANEL_INFO.CELL_REPAIR_JUDGE_GRADE = Left(strTemp, intPos - 1)
                                Case intCell_Line_Rescue_Flag_Pos:
                                    pPANEL_INFO.CELL_LINE_RESCUE_FLAG = Left(strTemp, intPos - 1)
                                Case intRework_Count_Pos:
                                    pPANEL_INFO.REPAIR_REWORK_COUNT = Left(strTemp, intPos - 1)
                                End Select
                                
                                strTemp = Mid(strTemp, intPos + 1)
                                intPos = InStr(strTemp, ",")
                            Wend
                            intIndex = intIndex + 1
                            Select Case intIndex
                            Case intCell_Repair_Judge_Grade_Pos:
                                pPANEL_INFO.CELL_REPAIR_JUDGE_GRADE = strTemp
                            Case intCell_Line_Rescue_Flag_Pos:
                                pPANEL_INFO.CELL_LINE_RESCUE_FLAG = strTemp
                            Case intRework_Count_Pos:
                                pPANEL_INFO.REPAIR_REWORK_COUNT = strTemp
                            End Select
                            bolRead_Data = True
                        End If
                        Line Input #intFileNum, strTemp
                    Wend
                End If
            ElseIf InStr(UCase(strTemp), "CACRP_REPAIR_DEFECT_DATA_BEGIN") > 0 Then
                bolRead_Data = False
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    intItem_Index = 0
                    While (intPos > 0) And (bolRead_Data = False)
                        intItem_Index = intItem_Index + 1
                        strTitle = Left(strTemp, intPos - 1)
                        Select Case UCase(strTitle)
                        Case "CARBONIZATION_FLAG":
                            intCarbonization_Flag_Pos = intItem_Index
                        Case "CARBONIZATION_GRADE":
                            intCarbonization_Grade_Pos = intItem_Index
                        Case "CARBONIZATION_REWORK_COUNT":
                            intCarbonization_Rework_Count_Pos = intItem_Index
                        End Select
                        
                        strTemp = Mid(strTemp, intPos + 1)
                        intPos = InStr(strTemp, ",")
                    Wend
                    Select Case UCase(strTemp)
                    Case "CARBONIZATION_FLAG":
                        intCarbonization_Flag_Pos = intItem_Index
                    Case "CARBONIZATION_GRADE":
                        intCarbonization_Grade_Pos = intItem_Index
                    Case "CARBONIZATION_REWORK_COUNT":
                        intCarbonization_Rework_Count_Pos = intItem_Index
                    End Select
                    Line Input #intFileNum, strTemp
                    While UCase(strTemp) <> "CACRP_REPAIR_DEFECT_DATA_BEGIN"
                        intPos = InStr(strTemp, ",")
                        If intPos > 0 Then
                            intIndex = 0
                            While intPos > 0
                                intIndex = intIndex + 1
                                Select Case intIndex
                                Case intCarbonization_Flag_Pos:
                                    pPANEL_INFO.CARBONIZATION_FLAG = Left(strTemp, intPos - 1)
                                Case intCarbonization_Grade_Pos:
                                    pPANEL_INFO.CARBONIZATION_GRADE = Left(strTemp, intPos - 1)
                                Case intCarbonization_Rework_Count_Pos:
                                    pPANEL_INFO.CARBONIZATION_REWORK_COUNT = Left(strTemp, intPos - 1)
                                End Select
                                
                                strTemp = Mid(strTemp, intPos + 1)
                                intPos = InStr(strTemp, ",")
                            Wend
                            intIndex = intIndex + 1
                            Select Case intIndex
                            Case intCarbonization_Flag_Pos:
                                pPANEL_INFO.CARBONIZATION_FLAG = strTemp
                            Case intCarbonization_Grade_Pos:
                                pPANEL_INFO.CARBONIZATION_GRADE = strTemp
                            Case intCarbonization_Rework_Count_Pos:
                                pPANEL_INFO.CARBONIZATION_REWORK_COUNT = strTemp
                            End Select
                            bolRead_Data = True
                        End If
                        Line Input #intFileNum, strTemp
                    Wend
                End If
            End If
        Wend

        Close intFileNum
        
    End If
    
End Sub
'Lucas 2012.02.09  Ver.0.9.4 -----Show address of source to the JPS
'==========================================Start


Public Sub Read_Source_Defect_File(pPANEL_INFO As PANEL_INFO_ELEMENTS, ByVal pLocalPath As String, ByVal pFileName As String)

    Dim strTemp                             As String
    Dim strTitle                            As String
    Dim intFileNum                          As Integer
    Dim intPos                              As Integer
    Dim intItem_Index                       As Integer
    Dim intIndex                            As Integer
    Dim bolRead_Data                        As Boolean
    Dim i                                   As Integer
    Dim typPANEL_INFO                       As PANEL_INFO_ELEMENTS
    

    If Dir(pLocalPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open pLocalPath & pFileName For Input As intFileNum
        
For i = 1 To 15
 frmJudge.Controls("text" & i).Text = ""
Next i

While Not EOF(intFileNum)
  Line Input #intFileNum, strTemp
    If InStr(UCase(strTemp), "DEFECT_DATA_BEGIN") > 0 Then
       Line Input #intFileNum, strTemp
       Line Input #intFileNum, strTemp
                      If Left(strTemp, 12) = pPANEL_INFO.PANELID Then
                          intPos = InStr(strTemp, ",")
                           If intPos > 0 Then
                            intIndex = 0
                            While intPos > 0
                                intIndex = intIndex + 1
                                Select Case intIndex
                                Case "3":
                                    frmJudge.Text1.Text = Left(strTemp, intPos - 1)
                                Case "8":
                                    frmJudge.Text2.Text = Left(strTemp, intPos - 1)
                                Case "9":
                                    frmJudge.Text3.Text = Left(strTemp, intPos - 1)
                                Case "10":
                                    frmJudge.Text4.Text = Left(strTemp, intPos - 1)
                                Case "11":
                                    frmJudge.Text5.Text = Left(strTemp, intPos - 1)
                                End Select
                                
                                strTemp = Mid(strTemp, intPos + 1)
                                intPos = InStr(strTemp, ",")
                            Wend
                            End If
                     End If
      Line Input #intFileNum, strTemp
                     If Left(strTemp, 12) = pPANEL_INFO.PANELID Then
                          intPos = InStr(strTemp, ",")
                           If intPos > 0 Then
                            intIndex = 0
                            While intPos > 0
                                intIndex = intIndex + 1
                                Select Case intIndex
                                Case "3":
                                    frmJudge.Text6.Text = Left(strTemp, intPos - 1)
                                Case "8":
                                    frmJudge.Text7.Text = Left(strTemp, intPos - 1)
                                Case "9":
                                    frmJudge.Text8.Text = Left(strTemp, intPos - 1)
                                Case "10":
                                    frmJudge.Text9.Text = Left(strTemp, intPos - 1)
                                Case "11":
                                    frmJudge.Text10.Text = Left(strTemp, intPos - 1)
                                End Select
                                
                                strTemp = Mid(strTemp, intPos + 1)
                                intPos = InStr(strTemp, ",")
                            Wend
                            End If
                    End If
      Line Input #intFileNum, strTemp
                       If Left(strTemp, 12) = pPANEL_INFO.PANELID Then
                          intPos = InStr(strTemp, ",")
                           If intPos > 0 Then
                            intIndex = 0
                            While intPos > 0
                                intIndex = intIndex + 1
                                Select Case intIndex
                                Case "3":
                                    frmJudge.Text11.Text = Left(strTemp, intPos - 1)
                                Case "8":
                                    frmJudge.Text12.Text = Left(strTemp, intPos - 1)
                                Case "9":
                                    frmJudge.Text13.Text = Left(strTemp, intPos - 1)
                                Case "10":
                                    frmJudge.Text14.Text = Left(strTemp, intPos - 1)
                                Case "11":
                                    frmJudge.Text15.Text = Left(strTemp, intPos - 1)
                                End Select
                                
                                strTemp = Mid(strTemp, intPos + 1)
                                intPos = InStr(strTemp, ",")
                            Wend
                            End If
                    End If
                  
     End If
  
   Wend
   
   Close #intFileNum

End If
    
End Sub

Public Sub Read_Notice_File()

    Dim NOTICE_DATA                 As PUBLIC_NOTICE
    
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    Dim strDate                     As String
    Dim strTime                     As String
    Dim strMessage                  As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    Dim intCount                    As Integer
    Dim intLength                   As Integer
    
    Dim bolNew_File                 As Boolean
    
    strPath = App.PATH & "\Env\Standard_Info\"
    strFileName = "Public notice.txt"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        bolNew_File = True
        strTime = Format(FileDateTime(strPath & strFileName), "HHMMSS")
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        intCount = 0
        While (Not EOF(intFileNum)) And (bolNew_File = True)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "*")
            If intPos > 0 Then
                If UCase(Left(strTemp, intPos - 1)) <> "DATE" Then
                    strDate = Left(strTemp, intPos - 1)
                    strMessage = Mid(strTemp, intPos + 1)
                    If (strDate & " " & strMessage <> ENV.Get_NOTICE_MESSAGE_by_Index(1)) And (bolNew_File = True) Then
                        intCount = intCount + 1
                        If (Len(strDate) + Len(strMessage) + 1) > 80 Then
                            intCount = intCount + 1
                        End If
                    Else
                        bolNew_File = False
                    End If
                End If
            End If
        Wend
        
        Close intFileNum
            
        If bolNew_File = True Then
            Call ENV.Set_Notice_Count(intCount)
            
            If Dir(strPath & "Notice.txt", vbNormal) <> "" Then
                Kill strPath & "Notice.txt"
            End If
            Name strPath & strFileName As strPath & "Notice.txt"
            strFileName = "Notice.txt"
            
            intFileNum = FreeFile
            Open strPath & strFileName For Input As intFileNum
            
            intCount = 0
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, "*")
                If intPos > 0 Then
                    If UCase(Left(strTemp, intPos - 1)) <> "DATE" Then
                        strDate = Left(strTemp, intPos - 1)
                        strMessage = Mid(strTemp, intPos + 1)
                        intCount = intCount + 1
                        If (Len(strDate) + Len(strMessage) + 1) > 80 Then
                            Call ENV.Set_Notice_Data(intCount, strDate, strDate & " " & Left(strMessage, 80), strTime)
                            intCount = intCount + 1
                            Call ENV.Set_Notice_Data(intCount, strDate, Mid(strMessage, 81), strTime)
                        Else
                            Call ENV.Set_Notice_Data(intCount, strDate, strDate & " " & strMessage, strTime)
                        End If
                    End If
                End If
            Wend
            
            Close intFileNum
        Else
            Kill strPath & strFileName
        End If
        
    End If
    
End Sub

Public Sub Read_PFCD_ADDRESS_DATA(ByVal pFileName As String)

    Dim typPFCD_ADDRESS_DATA        As PFCD_ADDRESS_STRUCTURE
    
    Dim strPath                     As String
    Dim strTemp                     As String
    Dim strPRODUCT_ID               As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    
    strPath = App.PATH & "\Env\Standard_Info\"
    If Dir(strPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & pFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                If Left(strTemp, intPos - 1) <> "PRODUCT_ID" Then
                    With typPFCD_ADDRESS_DATA
                        .PRODUCT_ID = Left(strTemp, intPos - 1)
                        If .PRODUCT_ID <> "" Then
                            strPRODUCT_ID = .PRODUCT_ID
                        End If
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .PANEL_NO = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .W = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .W = 0
                        End If
                        
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .L = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .L = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .B1 = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .B1 = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .B2 = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .B2 = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .XC = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .XC = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .YC = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .YC = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .XO = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .XO = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        If Left(strTemp, intPos - 1) <> "" Then
                            .YO = CDbl(Left(strTemp, intPos - 1))
                        Else
                            .YO = 0
                        End If
                    
                        intPos = InStr(strTemp, ",")
                        .ORIGIN_LOCATION = Left(strTemp, intPos - 1)
                        .SOURCE_DIRECTION = Mid(strTemp, intPos + 1)
                    End With
                    If typPFCD_ADDRESS_DATA.PRODUCT_ID = "" Then
                        typPFCD_ADDRESS_DATA.PRODUCT_ID = strPRODUCT_ID
                    End If
                    Call Set_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA)
                End If
            End If
        Wend
        
        Close intFileNum
    End If
    
End Sub

Public Sub Set_PFCD_ADDRESS_DATA(pPFCD_ADDRESS_DATA As PFCD_ADDRESS_STRUCTURE)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PFCD_ADDRESS WHERE "
        strQuery = strQuery & "PRODUCT_ID = '" & pPFCD_ADDRESS_DATA.PRODUCT_ID & "' AND "
        strQuery = strQuery & "PANEL_NO = '" & pPFCD_ADDRESS_DATA.PANEL_NO & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            
            With pPFCD_ADDRESS_DATA
                strQuery = "UPDATE PFCD_ADDRESS SET "
                strQuery = strQuery & "W = " & .W & ", "
                strQuery = strQuery & "L = " & .L & ", "
                strQuery = strQuery & "B1 = " & .B1 & ", "
                strQuery = strQuery & "B2 = " & .B2 & ", "
                strQuery = strQuery & "XC = " & .XC & ", "
                strQuery = strQuery & "YC = " & .YC & ", "
                strQuery = strQuery & "XO = " & .XO & ", "
                strQuery = strQuery & "YO = " & .YO & ", "
                strQuery = strQuery & "ORIGIN_LOCATION = '" & .ORIGIN_LOCATION & "', "
                strQuery = strQuery & "SOURCE_DIRECTION = '" & .SOURCE_DIRECTION & "' WHERE "
                strQuery = strQuery & "PRODUCT_ID = '" & .PRODUCT_ID & "' AND "
                strQuery = strQuery & "PANEL_NO = '" & .PANEL_NO & "'"
                
                dbMyDB.Execute (strQuery)
            End With
        Else
            lstRecord.Close
            
            With pPFCD_ADDRESS_DATA
                strQuery = "INSERT INTO PFCD_ADDRESS VALUES ("
                strQuery = strQuery & "'" & .PRODUCT_ID & "', "
                strQuery = strQuery & "'" & .PANEL_NO & "', "
                strQuery = strQuery & .W & ", "
                strQuery = strQuery & .L & ", "
                strQuery = strQuery & .B1 & ", "
                strQuery = strQuery & .B2 & ", "
                strQuery = strQuery & .XC & ", "
                strQuery = strQuery & .YC & ", "
                strQuery = strQuery & .XO & ", "
                strQuery = strQuery & .YO & ", "
                strQuery = strQuery & "'" & .ORIGIN_LOCATION & "', "
                strQuery = strQuery & "'" & .SOURCE_DIRECTION & "')"
                
                dbMyDB.Execute (strQuery)
            End With
        End If
        
        
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Get_PFCD_ADDRESS_DATA(pPFCD_ADDRESS_DATA As PFCD_ADDRESS_STRUCTURE, ByVal pPRODUCTID As String, ByVal pPANEL_NO As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    pPANEL_NO = Trim(Str$(CInt(pPANEL_NO)))
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
        strQuery = "SELECT * FROM PFCD_ADDRESS WHERE "
        strQuery = strQuery & "PRODUCT_ID = '" & pPRODUCTID & "' AND "
        strQuery = strQuery & "PANEL_NO = '" & pPANEL_NO & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            
            With pPFCD_ADDRESS_DATA
                .PRODUCT_ID = lstRecord.Fields("PRODUCT_ID")
                .PANEL_NO = lstRecord.Fields("PANEL_NO")
                .W = lstRecord.Fields("W") / 1000
                .L = lstRecord.Fields("L") / 1000
                .B1 = lstRecord.Fields("B1") / 1000
                .B2 = lstRecord.Fields("B2") / 1000
                .XC = lstRecord.Fields("XC") / 1000
                .YC = lstRecord.Fields("YC") / 1000
                .XO = lstRecord.Fields("XO")
                .YO = lstRecord.Fields("YO")
                .ORIGIN_LOCATION = lstRecord.Fields("ORIGIN_LOCATION")
                .SOURCE_DIRECTION = lstRecord.Fields("SOURCE_DIRECTION")
            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Sub

Public Function Add_Judge_History_Grid(ByVal pPanelID As String) As Integer

    Dim intRow          As Integer
    Dim intCol          As Integer
    Dim intRow_Index    As Integer
    
    With frmMain.flxJudge_History
        intRow = 0
        For intRow_Index = 1 To .Rows - 1
            If .TextMatrix(intRow_Index, 0) = pPanelID Then
                intRow = intRow_Index
            End If
        Next intRow_Index
        If intRow = 0 Then
            intRow = .Rows
            .AddItem pPanelID
            .RowHeight(intRow) = 350
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
            Next intCol
            If .Rows > 6 Then
                .TopRow = .Rows - 6
            End If
            
            Add_Judge_History_Grid = .Rows - 1
        Else
            Add_Judge_History_Grid = intRow
        End If
    End With

End Function

Public Sub Read_PFCD_DATA()

    Dim typPFCD_DATA                As PFCD_DATA
    
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    
    strPath = App.PATH & "\Env\Standard_Info\"
    strFileName = "PFCD.PID"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                If Left(strTemp, intPos - 1) <> "PID" Then
                    With typPFCD_DATA
                        .PFCD = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .X_PIXEL_LENGTH = Left(strTemp, intPos - 1)
                        pubPANEL_INFO.X_TOTAL_PIXEL = .X_PIXEL_LENGTH
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .Y_PIXEL_LENGTH = Left(strTemp, intPos - 1)
                        pubPANEL_INFO.Y_TOTAL_PIXEL = .Y_PIXEL_LENGTH
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .DATA = Left(strTemp, intPos - 1)
                        pubPANEL_INFO.X_ONE_PIXEL_LENGTH = .DATA
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .GATE = Left(strTemp, intPos - 1)
                        pubPANEL_INFO.Y_ONE_PIXEL_LENGTH = .GATE
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .CSTC = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .MAX_PANEL = Left(strTemp, intPos - 1)
                        .PANEL_TYPE = Mid(strTemp, intPos + 1)
                    End With
                End If
            End If
        Wend
        
        Close intFileNum
    End If
    
End Sub

Public Sub Insert_PFCD_DATA(pPFCD_DATA As PFCD_DATA)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    If pPFCD_DATA.PFCD <> "" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "Parameter.mdb"
        
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM PFCD_DATA WHERE "
            strQuery = strQuery & "PFCD = '" & pPFCD_DATA.PFCD & "'"
        
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.Close
                strQuery = "UPDATE PFCD_DATA SET "
                strQuery = strQuery & "X_PIXEL_LENGTH = '" & pPFCD_DATA.X_PIXEL_LENGTH & "', "
                strQuery = strQuery & "Y_PIXEL_LENGTH = '" & pPFCD_DATA.Y_PIXEL_LENGTH & "', "
                strQuery = strQuery & "DATA = '" & pPFCD_DATA.DATA & "', "
                strQuery = strQuery & "GATE = '" & pPFCD_DATA.GATE & "', "
                strQuery = strQuery & "CSTC = '" & pPFCD_DATA.CSTC & "', "
                strQuery = strQuery & "MAX_PANEL = '" & pPFCD_DATA.MAX_PANEL & "', "
                strQuery = strQuery & "PANEL_TYPE = '" & pPFCD_DATA.PANEL_TYPE & "' WHERE "
                strQuery = strQuery & "PFCD = '" & pPFCD_DATA.PFCD & "'"
                
                dbMyDB.Execute (strQuery)
            Else
                lstRecord.Close
                With pPFCD_DATA
                    strQuery = "INSERT INTO PFCD_DATA VALUES ("
                    strQuery = strQuery & "'" & .PFCD & "', "
                    strQuery = strQuery & "'" & .X_PIXEL_LENGTH & "', "
                    strQuery = strQuery & "'" & .Y_PIXEL_LENGTH & "', "
                    strQuery = strQuery & "'" & .DATA & "', "
                    strQuery = strQuery & "'" & .GATE & "', "
                    strQuery = strQuery & "'" & .CSTC & "', "
                    strQuery = strQuery & "'" & .MAX_PANEL & "', "
                    strQuery = strQuery & "'" & .PANEL_TYPE & "')"
                    
                    dbMyDB.Execute (strQuery)
                End With
            End If
            
            dbMyDB.Close
        End If
    End If
    
End Sub

Public Sub Get_PFCD_DATA(pPFCD_DATA As PFCD_DATA, ByVal pPFCD As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PFCD_DATA WHERE "
        strQuery = strQuery & "PFCD = '" & pPFCD & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With pPFCD_DATA
                .PFCD = pPFCD
                .X_PIXEL_LENGTH = lstRecord.Fields("X_PIXEL_LENGTH")
                .Y_PIXEL_LENGTH = lstRecord.Fields("Y_PIXEL_LENGTH")
                .DATA = lstRecord.Fields("DATA")
                .GATE = lstRecord.Fields("GATE")
                .CSTC = lstRecord.Fields("CSTC")
                .MAX_PANEL = lstRecord.Fields("MAX_PANEL")
                .PANEL_TYPE = lstRecord.Fields("PANEL_TYPE")
            End With
        End If
        
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Sub

Public Function Get_DEFECT_DATA_by_CODE(ByVal pDEFECT_CODE As String) As RANK_DATA_STRUCTURE

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
     '============Leo 2012.05.22 Add Rank Level Start
   Dim intRankLevel                 As Integer
     '============Leo 2012.05.22 Add Rank Level end
    strDB_Path = App.PATH & "\DB\"
               'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
    If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
            strDB_FileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"

    Else
        strDB_FileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    End If
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM RANK_DATA WHERE "
        Select Case pubCST_INFO.PROCESS_NUM
        Case "3000":
            strQuery = strQuery & "RANK_DIVISION = 'LOI-1' AND "
        Case "3650":
            strQuery = strQuery & "RANK_DIVISION = 'RLOI-1' AND "
        Case "4600":
            strQuery = strQuery & "RANK_DIVISION = 'LOI-2' AND "
        Case "4650":
            strQuery = strQuery & "RANK_DIVISION = 'RLOI-2' AND "
        End Select
        
        strQuery = strQuery & "DEFECT_CODE ='" & pDEFECT_CODE & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With Get_DEFECT_DATA_by_CODE
                .RANK_DIVISION = lstRecord.Fields("RANK_DIVISION")
                .DEFECT_CODE = lstRecord.Fields("DEFECT_CODE")
                .DEFECT_NAME = lstRecord.Fields("DEFECT_NAME")
                .DEFECT_DIVISION = lstRecord.Fields("DEFECT_DIVISION")
                .DEFECT_TYPE = lstRecord.Fields("DEFECT_TYPE")
                .JUDGE_OR_NOT = lstRecord.Fields("JUDGE_OR_NOT")
                .USE_XY = lstRecord.Fields("USE_XY")
                .DETAIL_DIVISION = lstRecord.Fields("DETAIL_DIVISION")
                .ACCUMULATION = lstRecord.Fields("ACCUMULATION")
                .ADDRESS_COUNT = lstRecord.Fields("ADDRESS_COUNT")
                     '============Leo 2012.05.22 Add Rank Level Start
                For intRankLevel = 0 To UBound(RankLevel)
                    .Rank(intRankLevel) = lstRecord.Fields("RANK_" & RankLevel(intRankLevel))
                Next intRankLevel
'                .RANK_Y = lstRecord.Fields("RANK_Y")
'                .RANK_L = lstRecord.Fields("RANK_L")
'                .RANK_K = lstRecord.Fields("RANK_K")
'                .RANK_C = lstRecord.Fields("RANK_C")
'                .RANK_S = lstRecord.Fields("RANK_S")
                     '============Leo 2012.05.22 Add Rank Level end
                .ODF = lstRecord.Fields("ODF")
                .PRIORITY = lstRecord.Fields("PRIORITY")
                .POP_UP = lstRecord.Fields("POP_UP")
            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Function

Public Sub Put_File_To_Host(ByVal pFileName As String, ByVal pSub_Path As String, ByVal pLocal_Path As String)

    Dim FTP_OBJECT              As New clsFTP
    
    Dim strRemote_Path          As String
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    If FTP_OBJECT.Init_FTP_Client = True Then
        Call FTP_OBJECT.Open_Session
        strRemote_Path = FTP_OBJECT.Get_Path(cFTP_HOST)
        If Right(strRemote_Path, 1) <> "\" Then
            strRemote_Path = strRemote_Path & "\"
        End If
        strRemote_Path = strRemote_Path & pSub_Path & "\"
        If FTP_OBJECT.FTP_Put_File(pFileName, strRemote_Path, pLocal_Path) = True Then
            Call SaveLog("Put_File_To_Host", pFileName & " upload success.")
        Else
            Call SaveLog("Put_File_To_Host", pFileName & " upload fail.")
        End If
    Else
        Call SaveLog("Put_File_To_Host", "FTP initialize fail.")
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Put_File_To_Host", ErrMsg)
    
End Sub

Public Sub Get_File_From_Host(ByVal pFileName As String, ByVal pSub_Path As String)

    Dim FTP_OBJECT              As New clsFTP

    Dim strRemote_Path          As String
    Dim strLocal_Path           As String

    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    If FTP_OBJECT.Init_FTP_Client = True Then
        Call FTP_OBJECT.Open_Session
        strRemote_Path = FTP_OBJECT.Get_Path(cFTP_HOST)
        If Right(strRemote_Path, 1) <> "\" Then
            strRemote_Path = strRemote_Path & "\"
        End If
        strRemote_Path = strRemote_Path & pSub_Path & "\"
        strLocal_Path = App.PATH & "\Env\Standard_Info\"
        If FTP_OBJECT.FTP_Get_File(pFileName, strRemote_Path, strLocal_Path) = True Then
            Call SaveLog("Get_File_From_Host", pFileName & " download success.")
        Else
            Call SaveLog("Get_File_From_Host", pFileName & " download fail.")
        End If
        Call FTP_OBJECT.Close_Session
        Call FTP_OBJECT.Disconnect_FTP_Client
    Else
        Call SaveLog("Get_File_From_Host", "FTP initialize fail.")
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Get_File_From_Host", ErrMsg)
    
End Sub

Public Sub Get_File_From_Host_by_Path(ByVal pPath As String, ByVal pLocal_Path As String, ByVal pFileName As String)

    Dim FTP_OBJECT              As New clsFTP

    Dim strRemote_Path          As String
    Dim strLocal_Path           As String

    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    If FTP_OBJECT.Init_FTP_Client = True Then
        Call FTP_OBJECT.Open_Session
        strRemote_Path = pPath
        If Right(strRemote_Path, 1) <> "\" Then
            strRemote_Path = strRemote_Path & "\"
        End If
        strLocal_Path = pLocal_Path
        If Right(strLocal_Path, 1) <> "\" Then
            strLocal_Path = strLocal_Path & "\"
        End If
        If FTP_OBJECT.FTP_Get_File(pFileName, strRemote_Path, strLocal_Path) = True Then
            Call SaveLog("Get_File_From_Host_by_Path", pFileName & " download success.")
        Else
            Call SaveLog("Get_File_From_Host", pFileName & " download fail.")
        End If
        Call FTP_OBJECT.Close_Session
        Call FTP_OBJECT.Disconnect_FTP_Client
    Else
        Call SaveLog("Get_File_From_Host", "FTP initialize fail.")
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Get_File_From_Host_by_Path", ErrMsg)
    
End Sub

Public Sub Power_On_PG()

    Dim intPortNo                       As Integer
    
    intPortNo = EQP.Get_PG_PortID
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QPPO")
    End If

End Sub

Public Sub Standard_Files_Download()

    Dim FTP_OBJ                             As New clsFTP
    
    Dim strRemote_Path                      As String
    Dim strLocalPath                        As String
    Dim strFilePath                         As String
    Dim strFileName                         As String
    
    Dim bolResult                           As Boolean
    
    strLocalPath = App.PATH & "\Env\Standard_Info\"
    strFilePath = App.PATH & "\Env\"
    strFileName = "File_List.txt"
    
    If FTP_OBJ.Init_FTP_Client = True Then
        strRemote_Path = FTP_OBJ.Get_Path(cFTP_HOST)
        If Right(strRemote_Path, 1) <> "\" Then
            strRemote_Path = strRemote_Path & "\"
        End If
 'Lucas Ver.0.9.16 2012.03.20===========================Change the Path of Get_File_From_List
            strRemote_Path = strRemote_Path & "Table\"
        bolResult = FTP_OBJ.FTP_Get_File_from_List(strRemote_Path, strLocalPath, strFilePath, strFileName)
        strFilePath = App.PATH & "\DB\"
        strFileName = "STANDARD_INFO.mdb"
        If Dir(strFilePath & strFileName, vbNormal) <> "" Then
            Kill strFilePath & strFileName
        End If
        FileCopy strFilePath & "STANDARD_INFO_Temp.mdb", strFilePath & strFileName
        Call Read_Control
        Call EQP.Set_Control_Data
        Call Read_TFT_CF_PanelID
        Call Read_Check_MES_DATA
        Call Read_Assign_Grade
        
        Call Read_PreJudgeGradeChange1
        Call Read_PreJudgeGradeChange2
        Call Read_PreJudgeGradeChange3
        Call Read_PostJudgeOtherRule1
        Call Read_PostJudgeOtherRule2
        Call Read_PostJudgeOtherRule3
        Call Read_PostJudgeGradeChange1
        Call Read_PostJudgeGradeChange2
        Call Read_CheckPanelIDChangeGrade
        Call Read_ChangeGrade
        Call Read_ChangeGradeByDefectCode
        Call Read_RepairPointTimes
        Call Read_FlagChangeGrade
        Call Read_SKChange
        Call Decode_Auto_Alarm
        Call RANK_OBJ.Reset_SK_SETTING
        Call Read_Notice_File
    End If
    
End Sub

Public Sub Set_Version_Data(pVERSION_DATA As VERSION_DATA)

    Dim FTP_OBJ                             As New clsFTP
    
    Dim strRemotePath                       As String
    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    
    Dim bolResult                           As Boolean
    
    With pVERSION_DATA
        strPath = App.PATH & "\Env\"
        strFileName = "Version_" & ENV.Get_JPS_Name & ".dat"
        intFileNum = FreeFile
        
        Open strPath & strFileName For Output As intFileNum
        
        strTemp = .MACHINE_ID
        strTemp = strTemp & "," & .JPS_VERSION
        strTemp = strTemp & "," & .EQ_VERSION
        strTemp = strTemp & "," & .JPS_NAME
        strTemp = strTemp & "," & .INSTALL_DAY
        strTemp = strTemp & "," & .USER
        strTemp = strTemp & "," & .JPS_SETUP_PATH
        strTemp = strTemp & "," & .JPS_LOG_PATH
        strTemp = strTemp & "," & .JPS_SERVER_PATH
        
        Print #intFileNum, strTemp
        
        Close intFileNum
    End With
    
    If FTP_OBJ.Init_FTP_Client = True Then       'FTP Object Initialize
        strRemotePath = FTP_OBJ.Get_Path(cFTP_DEFECT)
        If Right(strRemotePath, 1) <> "\" Then
        strRemotePath = strRemotePath & "\" & "EQ_Config\" & "JPS\" & "Version"
        End If
        Call FTP_OBJ.Open_Session                     'FTP Session Open
        bolResult = FTP_OBJ.FTP_Put_File(strFileName, strRemotePath, strPath)
        If bolResult = False Then
            Call SaveLog("Defect_File_Upload", strFileName & " upload fail. Remote path : " & strRemotePath)
        End If
        FTP_OBJ.Close_Session
        FTP_OBJ.Disconnect_FTP_Client
    Else
        Call SaveLog("Set_Version_Data", "FTP object initialize fail.")
    End If

End Sub

Public Sub Get_Version_Data(pVERSION_DATA As VERSION_DATA)

    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    Dim intPos                              As Integer
    
    With pVERSION_DATA
        strPath = App.PATH & "\Env\"
        strFileName = "Version_" & ENV.Get_JPS_Name & ".dat"
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        Line Input #intFileNum, strTemp
        
        intPos = InStr(strTemp, ",")
        If intPos > 0 Then
            .MACHINE_ID = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_VERSION = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .EQ_VERSION = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_NAME = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .INSTALL_DAY = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .USER = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_SETUP_PATH = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_LOG_PATH = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            .JPS_SERVER_PATH = strTemp
        End If
        
        Close intFileNum
    End With

End Sub

Public Sub Get_Version_Data_by_FileName(pVERSION_DATA As VERSION_DATA, ByVal pFileName As String)

    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    Dim intPos                              As Integer
    
    With pVERSION_DATA
        strPath = App.PATH & "\Env\"
        intFileNum = FreeFile
        
        Open strPath & pFileName For Input As intFileNum
        
        Line Input #intFileNum, strTemp
        
        intPos = InStr(strTemp, ",")
        If intPos > 0 Then
            .MACHINE_ID = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_VERSION = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .EQ_VERSION = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_NAME = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .INSTALL_DAY = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .USER = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_SETUP_PATH = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            intPos = InStr(strTemp, ",")
            .JPS_LOG_PATH = Left(strTemp, intPos - 1)
            strTemp = Mid(strTemp, intPos + 1)
            
            .JPS_SERVER_PATH = strTemp
        End If
        
        Close intFileNum
    End With

End Sub

Public Function Get_Defect_Name(ByVal pDEFECT_CODE As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    '=================================20120105. For CD08  Ver.0.9.2
    If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
          strDB_FileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    Else
       strDB_FileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    End If
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM RANK_DATA WHERE "
        strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_CODE & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            
            Get_Defect_Name = lstRecord.Fields("DEFECT_NAME")
        End If
        
        lstRecord.Close
        
        dbMyDB.Close
    End If

End Function

Public Function Get_Defect_Address_Count(ByVal pDEFECT_CODE As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
               'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
    
     If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
           strDB_FileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
     Else
       strDB_FileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
    End If
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM RANK_DATA WHERE "
        strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_CODE & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            
            Get_Defect_Address_Count = lstRecord.Fields("ADDRESS_COUNT")
        End If
        
        lstRecord.Close
        
        dbMyDB.Close
    End If

End Function

Public Function Get_Server_Path() As String

    Dim FTP_OBJ                     As New clsFTP
    
    If FTP_OBJ.Init_FTP_Client = True Then       'FTP Object Initialize
        Get_Server_Path = FTP_OBJ.Get_Path(cFTP_HOST)
    Else
        Get_Server_Path = ""
    End If
    
End Function

Public Sub Decode_Before_Block_Uncontact(ByVal pPortID As Integer)
    
    Dim dbMyDB                  As Database
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strDB_New_FileName      As String
    Dim strQuery                As String
    Dim strResponse             As String
    
    Dim intRow                  As Integer
        
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        Call RANK_OBJ.Set_END_TIME(Format(DATE, "YYYY/MM/DD") & "_" & Format(TIME, "HH:MM:SS"))
        
        If frmMain.flxAlign_PanelID.TextMatrix(1, 0) <> "" Then
            strDB_Path = App.PATH & "\DB\"
            strDB_FileName = "Result.mdb"
            
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
                
                strQuery = "UPDATE PANEL_DATA SET "
                strQuery = strQuery & "TACT_TIME=" & CLng(RANK_OBJ.Get_Tact_Time) & " WHERE "
                strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"
                
                dbMyDB.Execute (strQuery)
                
                dbMyDB.Close
            End If
    'Lucas Ver.0.9.5 2012.02.13------If mode is not the "Full Auto",JPS will write the data file.
    '======================================================Start
            If frmMain.flxEQ_Information.TextMatrix(3, 1) <> "Full Auto" Then
                Call Make_Defect_File
            End If
    '=======================================================End
            Call EQP.Set_DEFECT_UPLOAD(True)
        End If
        
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = ""
        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = ""
        frmMain.lblPre_Judge.Caption = ""
        frmMain.lblPost_Judge.Caption = ""
        frmMain.lblPre_Loss_Code.Caption = ""
        Call RANK_OBJ.Init_Class
        Call EQP.Set_QDAC_COMMAND("")
    End If
        
End Sub

Public Sub Insert_Pattern_Start(ByVal pDEFECT_CODE As String, ByVal pPTN_NAME As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strFilePath             As String
    Dim strFileName             As String
    Dim strQuery                As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & " '"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = True Then
            Call SaveLog("Insert_Pattern_Start", RANK_OBJ.Get_Current_KEYID & " panel result data does not exist.")
        Else
            lstRecord.MoveFirst
            strFilePath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        Set dbMyDB = Workspaces(0).OpenDatabase(strFilePath & strFileName)
        
        strQuery = "SELECT * FROM PATTERN_INSPECTION WHERE "
        strQuery = strQuery & "PATTERN_NAME='" & pPTN_NAME & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            
            strQuery = "UPDATE PATTERN_INSPECTION SET "
            strQuery = strQuery & "PATTERN_START=" & Timer & " WHERE "
            strQuery = strQuery & "PATTERN_NAME='" & pPTN_NAME & "'"
            
            Call dbMyDB.Execute(strQuery)
        Else
            lstRecord.Close
            
            strQuery = "INSERT INTO PATTERN_INSPECTION VALUES ("
            strQuery = strQuery & "'" & pPTN_NAME & "', "
            strQuery = strQuery & Timer & ", "
            strQuery = strQuery & "0, "
            strQuery = strQuery & "0)"
            
            Call dbMyDB.Execute(strQuery)
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub Insert_Pattern_End(ByVal pDEFECT_CODE As String, ByVal pPTN_NAME As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strFilePath             As String
    Dim strFileName             As String
    Dim strQuery                As String
    
    Dim lngStart                As Long
    Dim lngEnd                  As Long
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & " '"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = True Then
            Call SaveLog("Insert_Pattern_Start", RANK_OBJ.Get_Current_KEYID & " panel result data does not exist.")
        Else
            lstRecord.MoveFirst
            strFilePath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        Set dbMyDB = Workspaces(0).OpenDatabase(strFilePath & strFileName)
    
        strQuery = "SELECT * FROM PATTERN_INSPECTION WHERE "
        strQuery = strQuery & "PATTERN_NAME='" & pPTN_NAME & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            lngStart = lstRecord.Fields("PATTERN_START")
        End If
        lstRecord.Close
        
        lngEnd = Timer
        
        strQuery = "UPDATE PATTERN_INSPECTION SET "
        strQuery = strQuery & "PATTERN_END=" & lngEnd & ", "
        strQuery = strQuery & "INSPECTION_TIME=" & lngEnd - lngStart & " WHERE "
        strQuery = strQuery & "PATTERN_NAME='" & pPTN_NAME & "'"
        
        Call dbMyDB.Execute(strQuery)
        
        dbMyDB.Close
    End If
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - Move function position from modCATST_Sequence module to modSubroutine
'      for remove a YBBC reply delay
'
'==========================================================================================================

Public Function Decode_CATST_Before_Block_Contact(ByVal pCommand As String) As Integer

    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    Dim typJOB_INFO             As JOB_DATA_STRUCTURE
    Dim typSHARE_INFO           As SHARE_DATA_STRUCTURE
    Dim typPANEL_DATA           As PANEL_DATA
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strRemote_Path          As String
    Dim strLocal_Path           As String
    Dim strFileName             As String
    Dim strMode_State           As String
    Dim strDB_New_FileName      As String
    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strJobDataExistFlag     As String
    Dim strShareExistFlag       As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strJob_Info_Length      As String
    Dim strShare_Info_Length    As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
    Dim strPFCD                 As String
    Dim strOWNER                As String
    Dim strSUB_Command          As String
    Dim strMES_DATA_Command     As String
    Dim strJOB_DATA_Command     As String
    Dim strSHARE_DATA_Command   As String
    Dim strMsg                  As String
    
    Dim intRow                  As Integer
    Dim intIndex                As Integer
    
    Decode_CATST_Before_Block_Contact = 0
    
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        pCommand = Mid(pCommand, 5)
        strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
        
    'For panel ID change-----Lucas20111127
'        pubPANEL_INFO.AAAAAA = strPanelID
        
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        strMESDataExistFlag = Mid(pCommand, 1, (cSIZE_FLAG * 3))
        strJobDataExistFlag = Mid(strMESDataExistFlag, 2, 1)
        strShareExistFlag = Right(strMESDataExistFlag, 1)
        strMESDataExistFlag = Left(strMESDataExistFlag, 1)
        pCommand = Mid(pCommand, (cSIZE_FLAG * 3) + 1)
                    
        Select Case strMESDataExistFlag
        Case "E":               'MES DATA enable & data exist
            strCST_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strCST_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strCST_Info_Length))
            If Len(strSUB_Command) <> CInt(strCST_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                Decode_CATST_Before_Block_Contact = 1
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strCST_Info_Length) + 1)
            Call Decode_CST_Information_Elements(strSUB_Command, typCST_INFO)
            
            strPanel_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strMES_DATA_Command & strPanel_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strPanel_Info_Length))
            If Len(strSUB_Command) <> CInt(strPanel_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CATST_Before_Block_Contact = 0 Then
                    Decode_CATST_Before_Block_Contact = 2
                End If
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strPanel_Info_Length) + 1)
            Call Decode_PANEL_Information_Elements(strSUB_Command, typPANEL_INFO, typCST_INFO.PFCD)
                      
            strJob_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strJob_Info_Length))
            If Len(strSUB_Command) <> CInt(strJob_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CATST_Before_Block_Contact = 0 Then
                    Decode_CATST_Before_Block_Contact = 3
                End If
            End If
            strJOB_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strJob_Info_Length) + 1)
            Call Decode_JOB_Information_Elements(strSUB_Command, typJOB_INFO)
            
            strShare_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strShare_Info_Length))
            If Len(strSUB_Command) <> CInt(strShare_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CATST_Before_Block_Contact = 0 Then
                    Decode_CATST_Before_Block_Contact = 4
                End If
            End If
            strSHARE_DATA_Command = strSUB_Command
            Call Decode_Share_Information_Elements(strSUB_Command, typSHARE_INFO)
            
'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - If current operation mode is ON, JPS download pattern list file from file server otherwise
'      skip the download function and save QADC parameter
'
'
'  Start of modify
'
'==========================================================================================================

            If frmMain.flxEQ_Information.TextMatrix(3, 1) = "Operator" Then
                strRemote_Path = ENV.Get_Path_Data("PATTERN LIST")
                strLocal_Path = App.PATH & "\Env\Standard_Info\"
                strFileName = UCase(Left(typCST_INFO.OWNER, 1) & Mid(typCST_INFO.PFCD, 3, 5) & typCST_INFO.PROCESS_NUM) & ".csv"
'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - If changed PFCD or process number or pattern list file does not exist in local path, JPS download
'     pattern list file from File Server.
'
'==========================================================================================================
                If (Mid(typCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5)) Then
                    Call Get_File_From_Host(strFileName, "Pattern")
                ElseIf Dir(strLocal_Path & strFileName, vbNormal) = "" Then
                    Call Get_File_From_Host(strFileName, "Pattern")
                End If
                Call EQP.Read_PATTERN_LIST(strFileName)
                Call EQP.Set_PATTERN_LIST(strFileName)
            Else
                Call EQP.Set_MES_Data_for_API(strMESDataExistFlag, strJobDataExistFlag, strShareExistFlag, strMES_DATA_Command, strJOB_DATA_Command, strSHARE_DATA_Command)
            End If
            
'==========================================================================================================
'
'  End of modify
'
'==========================================================================================================

            With typPANEL_DATA
                .KEYID = typPANEL_INFO.PANELID & "_" & Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS")
                .TIME = Format(TIME, "HH:MM:SS")
                .PANELID = typPANEL_INFO.PANELID
                .PANEL_GRADE = Space(2)
                .PANEL_LOSSCODE = Space(5)
                .USER_NAME = frmMain.lblUser.Caption
                .RUN_DATE = CLng(Format(DATE, "YYYYMMDD"))
                .RUN_TIME = CLng(Format(TIME, "HHMMSS"))
                .TACT_TIME = 0
                .PATH = ""
                .FILENAME = ""
            End With
            If Make_CST_DATA_DB(Format(DATE, "MM"), Format(DATE, "DD"), typPANEL_DATA, typPANEL_INFO) = True Then
                Call SaveLog("Decode_CATST_Before_Block_Contact", typPANEL_INFO.PANELID & " data base create success.")
                Call Insert_Panel_MES_Data(typPANEL_DATA, typCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
                With frmMain.flxJudge_History
                    intRow = Add_Judge_History_Grid(typPANEL_DATA.PANELID)
                    .TextMatrix(intRow, 1) = frmMain.lblUser.Caption
                    .TextMatrix(intRow, 2) = typCST_INFO.PROCESS_NUM
                    .TextMatrix(intRow, 3) = ""
                    .TextMatrix(intRow, 4) = ""
                    .TextMatrix(intRow, 5) = ""
                    .TextMatrix(intRow, 6) = ""
    
                End With
            End If
            Call Save_MES_Data(typCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
            Call Get_MES_Data(pubCST_INFO, pubPANEL_INFO, pubJOB_INFO, pubSHARE_INFO)
            
            If Mid(typCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5) Then
                Call Get_File_From_Host("PFCD.PID", "Table")
                Call Read_PFCD_DATA
                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv", "Address")
                Call Read_PFCD_ADDRESS_DATA(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv")
                If frmMain.flxEQ_Information.TextMatrix(3, 1) = "Operator" Then
                    Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QSMY" & Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5) & frmMain.flxMES_Data.TextMatrix(3, 1))
                End If
'                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_address.csv")
            End If
                        
    '        Call Decode_PANEL_Information_Elements(pCommand, typPANEL_INFO)
            Call Set_MES_Data(pubCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
        
'                'TFT, CF Panel ID Check
'            strMsg = Check_TFT_CF_PanelID(typPANEL_INFO.PANELID)
'            If strMsg = "" Then
'                'Check MES Data
'                strMsg = Check_MES_Data(pubCST_INFO, typPANEL_INFO, typJOB_INFO)
'                If strMsg <> "" Then
'                    Call Show_Message("Abnormal MES Data", strMsg)
'                End If
'            Else
'                Call Show_Message("Abnormal Panel ID", strMsg)
'            End If
        Case "N":               'MES DATA enable & data not exist
        Case "D":               'MES DATA disable
        Case "S":               'In 1st inline light on
            strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
            pCommand = Mid(pCommand, cSIZE_PFCD + 1)
            strOWNER = Mid(pCommand, cSIZE_OWNER)
        End Select
        
        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = strPanelID
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = ""
    End If
    Call EQP.Set_RBBC_Command("")
    
End Function

'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - Move function position from modCALOI_Sequence module to modSubroutine
'      for remove a YBBC reply delay
'
'==========================================================================================================

Public Function Decode_CALOI_Before_Block_Contact(ByVal pCommand As String) As Integer

    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    Dim typJOB_INFO             As JOB_DATA_STRUCTURE
    Dim typSHARE_INFO           As SHARE_DATA_STRUCTURE
    Dim typPANEL_DATA           As PANEL_DATA
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strRemote_Path          As String
    Dim strLocal_Path           As String
    Dim strFileName             As String
    Dim strDB_New_FileName      As String
    Dim strSubCommand           As String
    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strJobDataExistFlag     As String
    Dim strShareExistFlag       As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strJob_Info_Length      As String
    Dim strShare_Info_Length    As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
    Dim strPFCD                 As String
    Dim strOWNER                As String
    Dim strSUB_Command          As String
    Dim strMES_DATA_Command     As String
    Dim strJOB_DATA_Command     As String
    Dim strSHARE_DATA_Command   As String
    Dim strMsg                  As String
    
    Dim intRow                  As Integer
    Dim intIndex                As Integer
    
    Decode_CALOI_Before_Block_Contact = 0
    
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        pCommand = Mid(pCommand, 5)
        strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        strMESDataExistFlag = Mid(pCommand, 1, (cSIZE_FLAG * 3))
        strJobDataExistFlag = Mid(strMESDataExistFlag, 2, 1)
        strShareExistFlag = Right(strMESDataExistFlag, 1)
        strMESDataExistFlag = Left(strMESDataExistFlag, 1)
        pCommand = Mid(pCommand, (cSIZE_FLAG * 3) + 1)
        strSubCommand = Right(pCommand, 24)
        strPanelType1 = Left(strSubCommand, 12)
        strPanelType2 = Right(strSubCommand, 12)
                    
        Select Case strMESDataExistFlag
        Case "E":               'MES DATA enable & data exist
            strCST_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strCST_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strCST_Info_Length))
            If Len(strSUB_Command) <> CInt(strCST_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                Decode_CALOI_Before_Block_Contact = 1
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strCST_Info_Length) + 1)
            Call Decode_CST_Information_Elements(strSUB_Command, typCST_INFO)
            
            
            strPanel_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strMES_DATA_Command & strPanel_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strPanel_Info_Length))
            If Len(strSUB_Command) <> CInt(strPanel_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CALOI_Before_Block_Contact = 0 Then
                    Decode_CALOI_Before_Block_Contact = 2
                End If
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strPanel_Info_Length) + 1)
            Call Decode_PANEL_Information_Elements(strSUB_Command, typPANEL_INFO, typCST_INFO.PFCD)
            
            strJob_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strJob_Info_Length))
            If Len(strSUB_Command) <> CInt(strJob_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CALOI_Before_Block_Contact = 0 Then
                    Decode_CALOI_Before_Block_Contact = 3
                End If
            End If
            strJOB_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strJob_Info_Length) + 1)
            Call Decode_JOB_Information_Elements(strSUB_Command, typJOB_INFO)
            
            strShare_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strShare_Info_Length))
            If Len(strSUB_Command) <> CInt(strShare_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_CALOI_Before_Block_Contact = 0 Then
                    Decode_CALOI_Before_Block_Contact = 4
                End If
            End If
            strSHARE_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strShare_Info_Length) + 1)
            Call Decode_Share_Information_Elements(strSUB_Command, typSHARE_INFO)
            
            Call EQP.Set_MES_Data_for_API(strMESDataExistFlag, strJobDataExistFlag, strShareExistFlag, strMES_DATA_Command, strJOB_DATA_Command, strSHARE_DATA_Command)
            
'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - If changed PFCD or process number or pattern list file does not exist in local path, JPS download
'     pattern list file from File Server.
'
'==========================================================================================================
            strRemote_Path = ENV.Get_Path_Data("PATTERN LIST")
            strLocal_Path = App.PATH & "\Env\Standard_Info\"
            strFileName = UCase(Left(typCST_INFO.OWNER, 1) & Mid(typCST_INFO.PFCD, 3, 5) & typCST_INFO.PROCESS_NUM) & ".csv"
            If (Mid(typCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5)) Or (typCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Then
                Call Get_File_From_Host(strFileName, "Pattern")
            ElseIf Dir(strLocal_Path & strFileName, vbNormal) = "" Then
                Call Get_File_From_Host(strFileName, "Pattern")
            End If
            Call EQP.Read_PATTERN_LIST(strFileName)
            Call EQP.Set_PATTERN_LIST(strFileName)
            
            With typPANEL_DATA
                .KEYID = typPANEL_INFO.PANELID & "_" & Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS")
                .TIME = Format(TIME, "HH:MM:SS")
                .PANELID = typPANEL_INFO.PANELID
                .PANEL_GRADE = Space(2)
                .PANEL_LOSSCODE = Space(5)
                .USER_NAME = frmMain.lblUser.Caption
                .RUN_DATE = CLng(Format(DATE, "YYYYMMDD"))
                .RUN_TIME = CLng(Format(TIME, "HHMMSS"))
                .TACT_TIME = 0
                .PANEL_TYPE1 = strPanelType1
                .PANEL_TYPE2 = strPanelType2
                .PATH = ""
                .FILENAME = ""
            End With
            If Make_CST_DATA_DB(Format(DATE, "MM"), Format(DATE, "DD"), typPANEL_DATA, typPANEL_INFO) = True Then
                Call SaveLog("Decode_CALOI_Before_Block_Contact", typPANEL_INFO.PANELID & " data base create success.")
                Call Insert_Panel_MES_Data(typPANEL_DATA, typCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)

                With frmMain.flxJudge_History
                    intRow = Add_Judge_History_Grid(typPANEL_DATA.PANELID)
                    While frmMain.lblUser.Caption = ""
                     Load frmLogin
                     frmLogin.Show
                    Wend
                    .TextMatrix(intRow, 1) = frmMain.lblUser.Caption
                    .TextMatrix(intRow, 2) = typCST_INFO.PROCESS_NUM
                    .TextMatrix(intRow, 3) = ""
                    .TextMatrix(intRow, 4) = ""
                    .TextMatrix(intRow, 5) = ""
                    .TextMatrix(intRow, 6) = ""
                End With
            End If
            Call Save_MES_Data(typCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
            Call Get_MES_Data(pubCST_INFO, pubPANEL_INFO, pubJOB_INFO, pubSHARE_INFO)
            
            If Mid(typCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5) Then
                Call Get_File_From_Host("PFCD.PID", "Table")
                Call Read_PFCD_DATA
                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv", "Address")
                Call Read_PFCD_ADDRESS_DATA(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv")
'                strDB_Path = App.PATH & "\DB\"
'                strDB_FileName = "STANDARD_INFO_Temp.mdb"
'                strDB_New_FileName = "STANDARD_INFO.mdb"
'
'                If Dir(strDB_Path & strDB_New_FileName, vbNormal) <> "" Then
'                    Kill strDB_Path & strDB_New_FileName
'                End If
'                FileCopy strDB_Path & strDB_FileName, strDB_Path & strDB_New_FileName
                If frmMain.flxEQ_Information.TextMatrix(3, 1) = "Operator" Then
                    Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QSMY" & Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5) & frmMain.flxMES_Data.TextMatrix(3, 1))
                End If
                
'                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_address.csv")
            End If
            
    '        Call Decode_PANEL_Information_Elements(pCommand, typPANEL_INFO)
'            Call Set_MES_Data(pubCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
'
'                'TFT, CF Panel ID Check
'            strMsg = Check_TFT_CF_PanelID(typPANEL_INFO.PANELID)
'            If strMsg = "" Then
'                'Check MES Data
'                strMsg = Check_MES_Data(pubCST_INFO, typPANEL_INFO, typJOB_INFO)
'                If strMsg <> "" Then
'                    Call Show_Message("Abnormal MES Data", strMsg)
'                End If
'            Else
'                Call Show_Message("Abnormal Panel ID", strMsg)
'            End If
        Case "N":               'MES DATA enable & data not exist
        Case "D":               'MES DATA disable
        Case "S":               'In 1st inline light on
            strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
            pCommand = Mid(pCommand, cSIZE_PFCD + 1)
            strOWNER = Mid(pCommand, cSIZE_OWNER)
        End Select
        
        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = strPanelID
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = ""
    End If
    Call EQP.Set_RBBC_Command("")
    
End Function

Public Sub Decode_CATST_After_Block_Contact(ByVal pCommand As String)

    Dim typDEFECT_DATA()        As DEFECT_DATA_STRUCTURE
    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    Dim typJOB_INFO             As JOB_DATA_STRUCTURE
    Dim typSHARE_INFO           As SHARE_DATA_STRUCTURE
    Dim typPANEL_DATA           As PANEL_DATA

    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
    Dim strPFCD                 As String
    Dim strOWNER                As String
    Dim strMsg                  As String
    Dim strStatus               As String
    Dim strCommand              As String
    Dim strMES_Exist            As String
    Dim strJOB_Exist            As String
    Dim strSHARE_Exist          As String
    Dim strMES_DATA             As String
    Dim strJOB_DATA             As String
    Dim strSHARE_DATA           As String
    Dim strGrade                As String
    Dim strMode_State           As String
    Dim strLocal_Path           As String
    Dim strFileName             As String
    
    Dim intPortNo               As Integer
    Dim intDefect_Count         As Integer
    Dim intIndex                As Integer
    Dim intRow                  As Integer
    
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        pCommand = Mid(pCommand, 5)
        strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        strMESDataExistFlag = Mid(pCommand, 1, cSIZE_FLAG)
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        Select Case strMESDataExistFlag
        Case "E":               'MES DATA enable & data exist
        Case "N":               'MES DATA enable & data not exist
        Case "D":               'MES DATA disable
        Case "S":               'In 1st inline light on
            strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
            pCommand = Mid(pCommand, cSIZE_PFCD + 1)
            strOWNER = Mid(pCommand, cSIZE_OWNER)
        End Select
        
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = strPanelID
        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = ""
        Call RANK_OBJ.Set_START_TIME(Format(DATE, "YYYY/MM/DD") & "_" & Format(TIME, "HH:MM:SS"))
        Select Case frmMain.flxEQ_Information.TextMatrix(3, 1)
        Case "Operator":
            strMode_State = "ON"
        Case "Auto and RJS":
            strMode_State = "IA"
        Case "Full Auto":
            strMode_State = "FA"
        Case "EQ Pass":
            strMode_State = "EP"
        End Select
        
'Lucas Ver0.9.19 2012.04.01===========================For Assign Grade before sending QDAC
    strGrade = Assign_Grade(pubCST_INFO, pubPANEL_INFO)
         If strGrade <> "" Then
            Call SaveLog("Decode_CATST_After_Block_Contact", "Assign Grade : " & strGrade)
            frmMain.lblPost_Judge.Caption = strGrade
            frmMain.flxMES_Data.TextMatrix(18, 1) = strGrade
            intRow = frmMain.flxJudge_History.Rows - 1
            frmMain.flxJudge_History.TextMatrix(intRow, 3) = strGrade
            frmMain.flxJudge_History.TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
            Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strGrade, frmMain.flxJudge_History.TextMatrix(intRow, 4), "")
         Else
'Lucas Ver0.9.19 2012.04.01===========================For Assign Grade before sending QDAC
        
           If strMode_State <> "ON" Then
            'Operator Mode
                Call ENV.Get_Device_Data_by_Name("API", intPortNo, strStatus)
                If intPortNo > 0 Then
                    'QDAC & Time & Panel ID & Owner & Process Number & PFCD & MES Data & Job Data
                    Call EQP.Get_MES_Data_for_API(strMES_Exist, strJOB_Exist, strSHARE_Exist, strMES_DATA, strJOB_DATA, strSHARE_DATA)
                    strCommand = Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS") & strPanelID & pubCST_INFO.OWNER & pubCST_INFO.PROCESS_NUM
                    strCommand = strCommand & pubCST_INFO.PFCD & strMES_DATA & "070" & strJOB_DATA & "207" & strSHARE_DATA
                    Call QUEUE.Put_Send_Command(intPortNo, "QDAC" & strCommand)
                    Call EQP.Set_QDAC_COMMAND("QDAC" & strCommand)
                End If
           Else    'Operator Mode
'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - If changed PFCD or process number or rank table file does not exist in local path, JPS download rank table.
'
'  Modify Date : 2011. 12. 28
'  Modify by K.H. KIM
'  Content
'    - Move rank table download and read location from Decode_Before_Block_Contact to Decode_After_Block_Contact
'    - Rank table database file name chage
'      If PFCD and Process Number are not change and rank table file already exist in local path, JPS not read
'      rank table data.
'
'==========================================================================================================
            If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
            Else
                strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
            End If
            strLocal_Path = App.PATH & "\Env\Standard_Info\"
                
            If (Mid(pubCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5)) Or (pubCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Or (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Then
                Call Get_File_From_Host(strFileName, strLocal_Path)
                If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                    strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                Else
                    strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                End If
                Call Read_Rank_Data(strFileName)
            ElseIf Dir(strLocal_Path & strFileName, vbNormal) = "" Then
                Call Get_File_From_Host(strFileName, strLocal_Path)
                           'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
               If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                    strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                Else
                    strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                End If
                Call Read_Rank_Data(strFileName)
            Else
                Call SaveLog("Decode_CATST_After_Block_Contact", "Rank DB create skip.")
                Call SaveLog("                                ", "Current PFCD : " & Mid(EQP.Get_Current_PFCD, 3, 5) & ", New PFCD : " & Mid(pubCST_INFO.PFCD, 3, 5))
                Call SaveLog("                                ", "Current PROC : " & EQP.Get_Current_PROCESSID & ", New PROC : " & pubCST_INFO.PROCESS_NUM)
            End If
            Load frmJudge
            Call EQP.Set_Current_PROCESSID(pubCST_INFO.PROCESS_NUM)
            Call EQP.Set_Current_PFCD(pubCST_INFO.PFCD)
            
'            strGrade = Assign_Grade(pubCST_INFO, pubPANEL_INFO)
'            If strGrade <> "" Then
'                Call SaveLog("Decode_CATST_After_Block_Contact", "Assign Grade : " & strGrade)
'                frmMain.lblPost_Judge.Caption = strGrade
'                frmMain.flxMES_Data.TextMatrix(18, 1) = strGrade
'                intRow = frmMain.flxJudge_History.Rows - 1
'                frmMain.flxJudge_History.TextMatrix(intRow, 3) = strGrade
'                frmMain.flxJudge_History.TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
'                Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strGrade, frmMain.flxJudge_History.TextMatrix(intRow, 4), "")
'            Else
                Call SaveLog("Decode_CATST_After_Block_Contact", "Manual Judge window load.")
                Call RANK_OBJ.Init_Class
                intDefect_Count = RANK_OBJ.Get_DEFECT_DATA_COUNT
                If intDefect_Count > 0 Then
                    ReDim typDEFECT_DATA(intDefect_Count)
                
                    For intIndex = 1 To intDefect_Count
                        With typDEFECT_DATA(intIndex)
                            If RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, _
                                                                 .GATE_ADDRESS, .GRADE, .RANK, .COLOR, .GRAY_LEVEL, .ACCUMULATION) = False Then
                                Call SaveLog("Decode_CATST_After_Block_Contact", "Defect Data loading fail. Index : " & intIndex)
                            End If
                        End With
                    Next intIndex
                End If
    '            Call Read_Rank_Data
    '
    '            Load frmJudge
                Call Power_On_PG
                frmJudge.Show
 'Lucas Ver0.9.29 2012.05.22---Show Alarm Msg after Block Contact
            'TFT, CF Panel ID Check
            strMsg = Check_TFT_CF_PanelID(pubPANEL_INFO.PANELID)
            If strMsg = "" Then
                'Check MES Data
                strMsg = Check_MES_Data(pubCST_INFO, pubPANEL_INFO, typJOB_INFO)
                If strMsg <> "" Then
                    Call Show_Message("Abnormal MES Data", strMsg)
                End If
            Else
                Call Show_Message("Abnormal Panel ID", strMsg)
            End If
                                
  'Lucas Ver0.9.29 2012.05.22---Show Alarm Msg after Block Contact
  
'            End If
        End If
      End If
'        intPortNo = EQP.Get_PG_PortID
'
'        If intPortNo > 0 Then
'            strCommand = "QSMY" & Mid(pubCST_INFO.PFCD, 3, 5)
'            Call QUEUE.Put_Send_Command(intPortNo, strCommand)
'        End If
   
    Else
            Call EQP.Set_Re_Contact_Flag(False)
            Call EQP.set_Re_Alignment_Flag(False)
    End If
    Call EQP.Set_RABC_Command("")
    
End Sub

Public Sub Decode_CALOI_After_Block_Contact(ByVal pCommand As String)

    Dim typDEFECT_DATA()        As DEFECT_DATA_STRUCTURE
    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    Dim typJOB_INFO             As JOB_DATA_STRUCTURE
    Dim typSHARE_INFO           As SHARE_DATA_STRUCTURE
    Dim typPANEL_DATA           As PANEL_DATA

    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
    Dim strPFCD                 As String
    Dim strOWNER                As String
    Dim strMsg                  As String
    Dim strStatus               As String
    Dim strCommand              As String
    Dim strMES_Exist            As String
    Dim strJOB_Exist            As String
    Dim strSHARE_Exist          As String
    Dim strMES_DATA             As String
    Dim strJOB_DATA             As String
    Dim strSHARE_DATA           As String
    Dim strGrade                As String
    Dim strMode_State           As String
    Dim strLocal_Path           As String
    Dim strFileName             As String
    
    Dim intPortNo               As Integer
    Dim intDefect_Count         As Integer
    Dim intIndex                As Integer
    Dim intRow                  As Integer
    
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        pCommand = Mid(pCommand, 5)
        strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
        pCommand = Mid(pCommand, cSIZE_PANELID + 1)
        strMESDataExistFlag = Mid(pCommand, 1, cSIZE_FLAG)
        pCommand = Mid(pCommand, cSIZE_FLAG + 1)
        Select Case strMESDataExistFlag
        Case "E":               'MES DATA enable & data exist
        Case "N":               'MES DATA enable & data not exist
        Case "D":               'MES DATA disable
        Case "S":               'In 1st inline light on
            strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
            pCommand = Mid(pCommand, cSIZE_PFCD + 1)
            strOWNER = Mid(pCommand, cSIZE_OWNER)
        End Select
        
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = strPanelID
        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = ""
        Call RANK_OBJ.Set_START_TIME(Format(DATE, "YYYY/MM/DD") & "_" & Format(TIME, "HH:MM:SS"))
        Select Case frmMain.flxEQ_Information.TextMatrix(3, 1)
        Case "Operator":
            strMode_State = "ON"
        Case "Auto and RJS":
            strMode_State = "IA"
        Case "Full Auto":
            strMode_State = "FA"
        Case "EQ Pass":
            strMode_State = "EP"
        End Select
        
        If strMode_State = "ON" Then
'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - If changed PFCD or process number or rank table file does not exist in local path, JPS download rank table.
'
'  Modify Date : 2011. 12. 28
'  Modify by K.H. KIM
'  Content
'    - Move rank table download and read location from Decode_Before_Block_Contact to Decode_After_Block_Contact
'    - Rank table database file name chage
'      If PFCD and Process Number are not change and rank table file already exist in local path, JPS not read
'      rank table data.
'
'==========================================================================================================
            'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
          If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
               strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
          Else
               strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
          End If
               strLocal_Path = App.PATH & "\Env\Standard_Info\"
            If (Mid(pubCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5)) Or (pubCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Or (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Then
                Call Get_File_From_Host(strFileName, strLocal_Path)
                           'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
                If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                       strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                Else
                    strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                End If
                Call Read_Rank_Data(strFileName)
            ElseIf Dir(strLocal_Path & strFileName, vbNormal) = "" Then
                Call Get_File_From_Host(strFileName, strLocal_Path)
                           'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
                If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                      strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                Else
                    strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                End If
                Call Read_Rank_Data(strFileName)
            Else
                Call SaveLog("Decode_CALOI_After_Block_Contact", "Rank DB create skip.")
                Call SaveLog("                                ", "Current PFCD : " & Mid(EQP.Get_Current_PFCD, 3, 5) & ", New PFCD : " & Mid(pubCST_INFO.PFCD, 3, 5))
                Call SaveLog("                                ", "Current PROC : " & EQP.Get_Current_PROCESSID & ", New PROC : " & pubCST_INFO.PROCESS_NUM)
            End If
            Load frmJudge
            Call EQP.Set_Current_PROCESSID(pubCST_INFO.PROCESS_NUM)
            Call EQP.Set_Current_PFCD(pubCST_INFO.PFCD)
            
            strGrade = Assign_Grade(pubCST_INFO, pubPANEL_INFO)
            If strGrade <> "" Then
                Call SaveLog("Decode_CALOI_After_Block_Contact", "Assign Grade : " & strGrade)
                frmMain.lblPost_Judge.Caption = strGrade
'                frmMain.flxMES_Data.TextMatrix(18, 1) = strGrade
                intRow = frmMain.flxJudge_History.Rows - 1
                frmMain.flxJudge_History.TextMatrix(intRow, 3) = strGrade
                frmMain.flxJudge_History.TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
                Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strGrade, frmMain.flxJudge_History.TextMatrix(intRow, 4), "")
            Else
                Call SaveLog("Decode_CALOI_After_Block_Contact", "Manual Judge window load.")
                Call RANK_OBJ.Init_Class
                intDefect_Count = RANK_OBJ.Get_DEFECT_DATA_COUNT
                If intDefect_Count > 0 Then
                    ReDim typDEFECT_DATA(intDefect_Count)
                
                    For intIndex = 1 To intDefect_Count
                        With typDEFECT_DATA(intIndex)
                            If RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                                 .GRADE, .RANK, .COLOR, .GRAY_LEVEL, .ACCUMULATION) = False Then
                                Call SaveLog("Decode_CALOI_After_Block_Contact", "Defect Data loading fail. Index : " & intIndex)
                            End If
                        End With
                    Next intIndex
                End If
    '            Call Read_Rank_Data
    '
    '            Load frmJudge
                Call Power_On_PG
                frmJudge.Show
'Lucas Ver0.9.29 2012.05.22---Show Alarm Msg after Block Contact
            'TFT, CF Panel ID Check
            strMsg = Check_TFT_CF_PanelID(pubPANEL_INFO.PANELID)
            If strMsg = "" Then
                'Check MES Data
                strMsg = Check_MES_Data(pubCST_INFO, pubPANEL_INFO, typJOB_INFO)
                If strMsg <> "" Then
                    Call Show_Message("Abnormal MES Data", strMsg)
                End If
            Else
                Call Show_Message("Abnormal Panel ID", strMsg)
            End If
                                
  'Lucas Ver0.9.29 2012.05.22---Show Alarm Msg after Block Contact
                
            End If
        End If
   
        
'        intPortNo = EQP.Get_PG_PortID
'        If intPortNo > 0 Then
'            strCommand = "QSMY" & Mid(pubCST_INFO.PFCD, 3, 5)
'            Call QUEUE.Put_Send_Command(intPortNo, strCommand)
'        End If
    Else
        Call Power_On_PG
        Call EQP.Set_Re_Contact_Flag(False)
        Call EQP.set_Re_Alignment_Flag(False)
    End If
    
    Call EQP.Set_RABC_Command("")
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Auto alarm database
'
'==========================================================================================================
Public Sub Decode_Auto_Alarm()

    Dim dbMyDB              As Database
    
    Dim typAUTO_ALARM       As AUTO_ALARM_DATA
    
    Dim strDB_Path          As String
    Dim strDB_FileName      As String
    Dim strPath             As String
    Dim strFileName         As String
    Dim strQuery            As String
    Dim strTemp             As String
    
    Dim lngExpiry_Date      As Long
    Dim lngExpiry_Time      As Long
    Dim intFileNum          As Integer
    Dim intPos              As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Auto_Alarm.mdb"
    strPath = App.PATH & "\Env\Standard_Info\"
    strFileName = "Auto alarm.csv"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Kill strDB_Path & strDB_FileName
        End If
        
        Call Make_AUTO_ALARM_DB
        
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, ",")
            '========================================================
            '
            ' 2012. 01. 24
            ' Modified by K.H. KIM
            ' Content
            '    - Skip then field's title line
            '
            '========================================================
            If UCase(Left(strTemp, intPos - 1)) <> "PROCESS NUM" Then
                With typAUTO_ALARM
                    .PROCESS_NUM = Left(strTemp, intPos - 1)
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    .PFCD = Left(strTemp, intPos - 1)
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    .DEFECT_CODE = Left(strTemp, intPos - 1)
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    .RANK = Left(strTemp, intPos - 1)
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    .COUNT_TIME = CInt(Left(strTemp, intPos - 1))
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    intPos = InStr(strTemp, ",")
                    .COUNT = CInt(Left(strTemp, intPos - 1))
                    strTemp = Mid(strTemp, intPos + 1)
                    
                    .ALARM_TEXT = strTemp
                    
                    .EXPIRY_DATE = CLng(Format(DATE + ((1 / 24 / 60) * .COUNT_TIME), "YYYYMMDD"))
                    .EXPIRY_TIME = CLng(Format(TIME + ((1 / 24 / 60) * .COUNT_TIME), "HHMMSS"))
                    
                    strQuery = "INSERT INTO AUTO_ALARM_DATA VALUES ("
                    strQuery = strQuery & "'" & .PROCESS_NUM & "', "
                    strQuery = strQuery & "'" & .PFCD & "', "
                    strQuery = strQuery & "'" & .DEFECT_CODE & "', "
                    strQuery = strQuery & "'" & .RANK & "', "
                    strQuery = strQuery & .COUNT_TIME & ", "
                    strQuery = strQuery & .COUNT & ", "
                    strQuery = strQuery & "'" & .ALARM_TEXT & "', "
                    strQuery = strQuery & "0, "
                    strQuery = strQuery & .EXPIRY_DATE & ", "
                    strQuery = strQuery & .EXPIRY_TIME & ")"
                    
                    dbMyDB.Execute strQuery
                End With
            End If
        Wend
        
        Close intFileNum
    Else
        Call SaveLog("Decode_Auto_Alarm", strPath & strFileName & " does not exist.")
    End If
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Check auto alarm case
'
'  Modify Date : 2012. 03. 26
'  Modify by K.H. KIM
'  Content
'    - Delete parameter
'
'==========================================================================================================
Public Sub Check_Auto_Alarm()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim typAUTO_ALARM               As AUTO_ALARM_DATA
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intIndex                    As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Auto_Alarm.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        For intIndex = 1 To pubDefect_Count
            strQuery = "SELECT * FROM AUTO_ALARM_DATA WHERE "
            strQuery = strQuery & "PROCESS_NUM='" & pubCST_INFO.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PFCD='" & Mid(pubCST_INFO.PFCD, 3, 5) & "' AND "
'            strQuery = strQuery & "RANK='" & pubDEFECT_DATA(intIndex).RANK & "' AND "
            strQuery = strQuery & "DEFECT_CODE='" & pubDEFECT_DATA(intIndex).DEFECT_CODE & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                
                With typAUTO_ALARM
                    .PROCESS_NUM = lstRecord.Fields("PROCESS_NUM")
                    .PFCD = lstRecord.Fields("PFCD")
                    .DEFECT_CODE = lstRecord.Fields("DEFECT_CODE")
                    .RANK = lstRecord.Fields("RANK")
                    .COUNT_TIME = lstRecord.Fields("COUNT_TIME")
                    .COUNT = lstRecord.Fields("COUNT")
                    .ALARM_TEXT = lstRecord.Fields("ALARM_TEXT")
                    .CURRENT_COUNT = lstRecord.Fields("CURRENT_COUNT")
                    .EXPIRY_DATE = lstRecord.Fields("EXPIRY_DATE")
                    .EXPIRY_TIME = lstRecord.Fields("EXPIRY_TIME")
                    
                    lstRecord.Close
                    
                    .CURRENT_COUNT = .CURRENT_COUNT + 1
                    If .CURRENT_COUNT = .COUNT Then
                        Load frmAuto_Alarm
'                        frmAuto_Alarm.lblTitle.Caption = "Common defect occurred. DEFECT CODE : " & .DEFECT_CODE & ", RANK : " & .RANK
                        frmAuto_Alarm.lblTitle.Caption = .ALARM_TEXT
                        frmAuto_Alarm.Show
                        
                        strQuery = "UPDATE AUTO_ALARM_DATA SET "
                        strQuery = strQuery & "CURRENT_COUNT=" & 0 & " WHERE "
                        strQuery = strQuery & "PROCESS_NUM='" & pubCST_INFO.PROCESS_NUM & "' AND "
                        strQuery = strQuery & "PFCD='" & Mid(pubCST_INFO.PFCD, 3, 5) & "' AND "
'                        strQuery = strQuery & "RANK='" & pubDEFECT_DATA(intIndex).RANK & "' AND "
                        strQuery = strQuery & "DEFECT_CODE='" & pubDEFECT_DATA(intIndex).DEFECT_CODE & "'"
                        
                        dbMyDB.Execute (strQuery)
                        Call SaveLog("Check_Auto_Alarm", pubDEFECT_DATA(intIndex).DEFECT_CODE & " auto alarm occurred. Defect Count : " & .CURRENT_COUNT)
                    Else
                        strQuery = "UPDATE AUTO_ALARM_DATA SET "
                        strQuery = strQuery & "CURRENT_COUNT=" & .CURRENT_COUNT & " WHERE "
                        strQuery = strQuery & "PROCESS_NUM='" & pubCST_INFO.PROCESS_NUM & "' AND "
                        strQuery = strQuery & "PFCD='" & Mid(pubCST_INFO.PFCD, 3, 5) & "' AND "
'                        strQuery = strQuery & "RANK='" & pubDEFECT_DATA(intIndex).RANK & "' AND "
                        strQuery = strQuery & "DEFECT_CODE='" & pubDEFECT_DATA(intIndex).DEFECT_CODE & "'"
                        
                        dbMyDB.Execute (strQuery)
                        Call SaveLog("Check_Auto_Alarm", pubDEFECT_DATA(intIndex).DEFECT_CODE & " current count : " & .CURRENT_COUNT)
                    End If
                End With
            End If
        Next intIndex
        
        dbMyDB.Close
    End If
    
End Sub
