Attribute VB_Name = "modAPI_Sequence"
Option Explicit

Public Sub API_Sequence(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strReply                    As String
    Dim strDevice_State             As String
    
    Dim intPortID                   As Integer
    
    If pPortID <> 9 Then
        Call ENV.Get_Device_Data_by_Name("API", intPortID, strDevice_State)
    End If
    
    If (strDevice_State = cDEVICE_ONLINE) Or (UCase(Left(pCommand, 4)) = "PONA") Or (pPortID = 9) Then
        Select Case UCase(Left(pCommand, 4))
        Case "PONA":        'On line reply from API
            Call Decode_Online_API(pPortID, pCommand)
    '        Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "POFA":        'Off line reply from API
            Call Decode_Offline_API(pPortID, pCommand)
            Unload frmOffLine_Request
    '        Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "RFPI":        'Finish Panel Inspection command report from API
'            Call Decode_Finished_Panel_Inspection(pPortID, pCommand)
'Lucas 2012.02.15 Ver.0.9.7 ---------------For QSPO-->QJPG
            Call QUEUE.Put_Send_Command(pPortID, "YFPI")
            If pPortID > 0 Then
            Call ENV.Get_Device_Data_by_Name("CATST", intPortID, strDevice_State)
            Call Receive_Judge_Panel_Grade_Reply(intPortID, pCommand)
            End If
'END
            Call QUEUE.Put_Send_Command(pPortID, "QBLV")
        Case "PSMY":        'Setting Modify command reply
        Case "RBAM":        'Buzz and Message command from API
            strReply = Decode_Buzz_Message(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "PDAC":        'Do action after contact command
            Call Decode_Do_Action_After_Contact(pPortID, pCommand)
        Case "RPDD":        'Panel defect data command
            strReply = Decode_Panel_Defect_Data(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "RRAD":        'Repeat address defect command
            strReply = Decode_Repeat_Address_Defect(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "RAQP":        'API Q Panel command
            strReply = Decode_Q_Panel_Command(ByVal pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "PBLV":        'Back light value
            Call Decode_Backlight_Value(pPortID, pCommand)
        Case "RAPG":        'API Panel grade
            strReply = Decode_Panel_Grade(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "RPNA":        'Panel need Re-Alignment command
            Call Decode_Need_Realignment(pPortID, pCommand)
        Case "PDAA":        'Do action after Re-Alignment
        Case "REQS":        'EQ Status
            strReply = Decode_EQ_State_Report(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "PEQS":        'EQ Status
            Call Decode_EQ_State_Reply(pPortID, pCommand)
        End Select
    End If
    
End Sub

Private Sub Decode_Online_API(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typVERSION_DATA             As VERSION_DATA
    
    Dim strAPIVersion               As String
    
    strAPIVersion = Mid(pCommand, 5)
        
    Call State_Change(pPortID, "API", cDEVICE_ONLINE)

    With frmMain.flxAPI_Information
        .TextMatrix(1, 1) = frmMain.flxEQ_Information.TextMatrix(1, 1)
        .TextMatrix(2, 1) = frmMain.flxEQ_Information.TextMatrix(2, 1)
'        .TextMatrix(3, 1) = frmMain.flxEQ_Information.TextMatrix(3, 1)
        .TextMatrix(5, 1) = "CAAPI" & Mid(strAPIVersion, 4)
        Call ENV.Set_API_Version(.TextMatrix(5, 1))
    End With
    Call Get_Version_Data(typVERSION_DATA)
    typVERSION_DATA.EQ_VERSION = ENV.Get_API_Version
    Call Set_Version_Data(typVERSION_DATA)
    
End Sub

Private Sub Decode_Offline_API(ByVal pPortID As Integer, ByVal pCommand As String)

    Call State_Change(pPortID, "API", cDEVICE_OFFLINE)
    
End Sub

Private Sub Decode_Finished_Panel_Inspection(ByVal pPortID As Integer, ByVal pCommand As String)

    'Next EQ Sequence
    Dim intRow                  As Integer
    
'    With frmMain
'        .flxAlign_PanelID.TextMatrix(1, 0) = ""
'        .lblPre_Judge.Caption = ""
'        .lblPost_Judge.Caption = ""
'        .lblPre_Loss_Code.Caption = ""
'        For intRow = 0 To .flxMES_Data.Rows - 1
'            .flxMES_Data.TextMatrix(intRow, 1) = ""
'        Next intRow
'    End With
    
End Sub

Private Function Decode_Buzz_Message(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim strALCD                 As String
    Dim strALTEXT               As String
    Dim strStatus               As String
    
    Dim intLength               As Integer
    Dim intPortNo               As Integer
    
    pCommand = Mid(pCommand, 5)
    intLength = Len(pCommand)
    strALCD = Left(pCommand, cSIZE_ALARM_CODE)
    strALTEXT = Mid(pCommand, 5)
    
    Call Show_Message("Alarm from API", strALCD & " : " & strALTEXT)
    
    Decode_Buzz_Message = "YBAM"
'Lucas Ver.0.9.28 2012.05.16 PG alarm Didn't send to EQ
    
'    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
'
'    If intPortNo > 0 Then
'        Call QUEUE.Put_Send_Command(intPortNo, "QBAM" & strALCD & strALTEXT)
'    End If
'Lucas Ver.0.9.28 2012.05.16 PG alarm Didn't send to EQ
End Function

Private Function Make_FTP_Folder(ByVal pRemotePath As String) As Boolean

    Dim FTP_OBJ                 As New clsFTP
    
    Dim intResult               As Integer
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

    If FTP_OBJ.Init_FTP_Client = True Then       'FTP Object Initialize
        If Right(pRemotePath, 1) <> "\" Then
            pRemotePath = pRemotePath & "\"
        End If
        Call FTP_OBJ.Open_Session                     'FTP Session Open
        intResult = FTP_OBJ.Remote_Change_Directory(pRemotePath)
        If intResult = 0 Then
            Make_FTP_Folder = True
        Else
            Make_FTP_Folder = False
        End If
        FTP_OBJ.Close_Session
        FTP_OBJ.Disconnect_FTP_Client
    Else
        Make_FTP_Folder = False
    End If
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("FTP_Upload", ErrMsg)
    Make_FTP_Folder = False

End Function

Private Sub Decode_Do_Action_After_Contact(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strResponse             As String
    Dim strRemote_Path          As String
    Dim typHEADER_DATA          As DEFECT_FILE_HEADER
    Dim typPANEL_DATA           As DEFECT_FILE_PANEL_DATA
    
    strResponse = Mid(pCommand, 5, 1)
    If strResponse = "0" Then
        'API save image file to IBW server
    Else
        'API don't save image file to IBW server
'        Call Show_Message("File save fail", "API don't save image file to IBW server.")
        
    End If
    
    
    'Lucas---For CAAPI File Creat Ver.0.7.33 2011.12.26
    
    With typHEADER_DATA
        .JPS_VERSION = App.Major & "." & App.Minor & "." & App.Revision
        .FILE_CREATE_TIME = Format(DATE, "YYYY/MM/DD") & "_" & Format(TIME, "HH:MM:SS")
        .EQUIP_TYPE = Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5)
        .EQ_ID = frmMain.flxEQ_Information.TextMatrix(5, 1)
        .SUBEQ_ID = .EQ_ID
    End With
    With typPANEL_DATA
        .PANELID = Trim(pubPANEL_INFO.PANELID)
        .GLASS_TYPE = Trim(pubCST_INFO.OWNER)
        .PRODUCT_ID = Trim(pubPANEL_INFO.PRODUCTID)
        .PROCESS_ID = Trim(pubCST_INFO.PROCESS_NUM)
        .RECIPE_ID = Trim(pubSHARE_INFO.RECIPEID)
        .SALEORDER = Trim(pubSHARE_INFO.SALE_ORDER)
        .CF_GLASS_ID = Trim(pubSHARE_INFO.CF_GLASSID)
        .ARRAY_LOT_ID = Trim(pubSHARE_INFO.ARRAY_LOTID)
        .ARRAY_GLASS_ID = Trim(pubSHARE_INFO.ARRAY_GLASSID)
        .CF_GLASS_OX_INFO = Trim(pubSHARE_INFO.CF_GLASS_INFO)
        .TFT_PANEL_JUDGE = Trim(pubSHARE_INFO.TFT_PANEL_JUDGE)
        .GROUP_ID = Trim(pubPANEL_INFO.GROUP_ID)
    End With

    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Source\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create fail.")
    End If
    
    strRemote_Path = "Link\" & "CATST" & "\" & Mid(typPANEL_DATA.PRODUCT_ID, 3, 5) & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create fail.")
    End If



    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Image\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create fail.")
    End If
    
    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Error\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create fail.")
    End If
    
    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Backup\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Decode_Do_Action_After_Contact", strRemote_Path & " folder create fail.")
    End If

    
End Sub
Private Function Decode_Repeat_Address_Defect(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
    Dim typPATTERN_LIST                 As PATTERN_LIST_DATA
    
    Dim strDATA_ADDRESS(1 To 3)         As String
    Dim strGATE_ADDRESS(1 To 3)         As String
        
    Dim strDefect_Code                  As String
    
    Dim intDefect_Count                 As Integer
    Dim intIndex                        As Integer
    Dim intCol                          As Integer
    Dim intRow                          As Integer
    Dim intGrade_Count                  As Integer
    
    '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end
    Call Reset_Interlock
    
    intRow = frmJudge.flxDefect_List.Rows - 1
    strDefect_Code = frmJudge.flxDefect_List.TextMatrix(intRow, 0)
    
    pCommand = Mid(pCommand, 5)
    
    For intIndex = 1 To 3
        strDATA_ADDRESS(intIndex) = Left(pCommand, 5)
        strGATE_ADDRESS(intIndex) = Mid(pCommand, 6, 5)
        frmJudge.txtX_Data(intIndex - 1).Text = strDATA_ADDRESS(intIndex)
        frmJudge.txtY_Gate(intIndex - 1).Text = strGATE_ADDRESS(intIndex)
        If intIndex < 3 Then
            pCommand = Mid(pCommand, 11)
        End If
    Next intIndex
    
    intIndex = 0
    For intCol = 2 To 7 Step 2
        intIndex = intIndex + 1
        With frmJudge.flxDefect_List
            .TextMatrix(intRow, intCol) = strDATA_ADDRESS(intIndex)
            .TextMatrix(intRow, intCol + 1) = strGATE_ADDRESS(intIndex)
        End With
    Next intCol
    
    'Check Defect Type
    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, strDefect_Code, intGrade_Count)

    If typRANK_DATA.DEFECT_TYPE = "R" Then
        'Manual judge
'        If typRANK_DATA.POP_UP = "E" Then
            'Sub window pop up
            Load frmManual_Judge
            
            With frmManual_Judge
            '============Leo 2012.05.22 Add Rank Level Start
                For intRankLevel = 0 To UBound(RankLevel)
                    If (Trim(typRANK_DATA.RANK(intRankLevel)) <> "0") And (Trim(typRANK_DATA.RANK(intRankLevel)) <> "-") Then
                        .lblGrade(intRankLevel).Caption = RankLevel(intRankLevel)
                        .optSpec_Value(intRankLevel).Caption = typRANK_DATA.RANK(intRankLevel)
                        .lblGrade(intRankLevel).Visible = True
                        .optSpec_Value(intRankLevel).Visible = True
                    End If
                Next intRankLevel
                                        
'                If (Trim(typRANK_DATA.RANK_Y) <> "0") And (Trim(typRANK_DATA.RANK_Y) <> "-") Then
'                    .lblGrade(0).Caption = "Y"
'                    .optSpec_Value(0).Caption = typRANK_DATA.RANK_Y
'                    .lblGrade(0).Visible = True
'                    .optSpec_Value(0).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_L) <> "0") And (Trim(typRANK_DATA.RANK_L) <> "-") Then
'                    .lblGrade(1).Caption = "L"
'                    .optSpec_Value(1).Caption = typRANK_DATA.RANK_L
'                    .lblGrade(1).Visible = True
'                    .optSpec_Value(1).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_K) <> "0") And (Trim(typRANK_DATA.RANK_K) <> "-") Then
'                    .lblGrade(2).Caption = "K"
'                    .optSpec_Value(2).Caption = typRANK_DATA.RANK_K
'                    .lblGrade(2).Visible = True
'                    .optSpec_Value(2).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_C) <> "0") And (Trim(typRANK_DATA.RANK_C) <> "-") Then
'                    .lblGrade(3).Caption = "C"
'                    .optSpec_Value(3).Caption = typRANK_DATA.RANK_C
'                    .lblGrade(3).Visible = True
'                    .optSpec_Value(3).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_S) <> "0") And (Trim(typRANK_DATA.RANK_S) <> "-") Then
'                    .lblGrade(4).Caption = "S"
'                    .optSpec_Value(4).Caption = typRANK_DATA.RANK_S
'                    .lblGrade(4).Visible = True
'                    .optSpec_Value(4).Visible = True
'                End If
 '============Leo 2012.05.22 Add Rank Level End
              
                .lblDefect_Code.Caption = strDefect_Code
                .lblDefect_Name.Text = frmJudge.flxDefect_List.TextMatrix(intRow, 1)
                .lstData_Address.Clear
                .lstGate_Address.Clear
                
                If Mid(.lblDefect_Code.Caption, 2, 1) = "M" Then
                    For intIndex = 1 To 3
                        .lstData_Address.AddItem strDATA_ADDRESS(intIndex)
                        .lstGate_Address.AddItem strGATE_ADDRESS(intIndex)
                    Next intIndex
                Else
                    .lstData_Address.AddItem strDATA_ADDRESS(1)
                    .lstGate_Address.AddItem strGATE_ADDRESS(1)
                End If
            End With
            
            frmManual_Judge.Show
'        End If
    End If
    frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 11) = typRANK_DATA.DETAIL_DIVISION
    intIndex = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    
    If intIndex > 0 Then
        With typPATTERN_LIST
            Call EQP.Get_PATTERN_LIST_by_Index(intIndex, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 12) = .PATTERN_NAME
        frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 13) = .LEVEL
        End With
    End If
    
'    Dim strAddress()            As String
'
'    Dim strDefectCode           As String
'    Dim strTemp                 As String
'
'    Dim intDataCount            As Integer
'    Dim intDataIndex            As Integer
'
'    strTemp = Mid(pCommand, 5)
'
'    strDefectCode = Mid(pCommand, 1, 5)
'    strTemp = Mid(pCommand, 6)
'
'    intDataCount = Len(strTemp) / 5
'    If intDataCount > 0 Then
'        ReDim strAddress(intDataCount)
'
'        For intDataIndex = 1 To intDataCount
'            strAddress(intDataIndex) = Mid(strTemp, 1, 5)
'            If Len(strTemp) > 5 Then
'                strTemp = Mid(strTemp, 6)
'            End If
'        Next intDataIndex
'    End If
        
    Decode_Repeat_Address_Defect = "YRAD"
    
End Function

Private Function Decode_Q_Panel_Command(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim strPanelID              As String
    
    strPanelID = Mid(pCommand, 5)
    
    Decode_Q_Panel_Command = "YAQP"
    
End Function

Private Sub Decode_Backlight_Value(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strValue                As String
    
    strValue = Mid(pCommand, 5)
    
    Call EQP.Set_BackLight_Value(strValue)
    
End Sub

Private Sub Decode_Need_Realignment(ByVal pPortID As Integer, ByVal pCommand)

    Dim intPortNo               As Integer
    
    Dim strStatus               As String
    
'    Call QUEUE.Put_Send_Command(pPortID, "QDAA")
    Call QUEUE.Put_Send_Command(pPortID, "YPNA")
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QRCO")
    End If
    
End Sub

Private Function Decode_Panel_Grade(ByVal pPortID As Integer, ByVal pCommand) As String

    Dim dbMyDB                              As Database
    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE

    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    Dim typGRADE_DEFECT_DATA                As DEFECT_DATA_STRUCTURE
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strPanelID              As String
    Dim strGrade                As String
    Dim strLossCode             As String
    Dim strStatus               As String
    Dim strBackLight_Value      As String
    
    Dim intPortID               As Integer
    Dim intSpace                As Integer
    Dim intRow                  As Integer
    Dim intGrade_Defect_Index   As Integer
    Dim intDefect_Count         As Integer
    Dim strPoint_Defect_Rank    As String
    pCommand = Mid(pCommand, 5)
    strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
    pCommand = Mid(pCommand, cSIZE_PANELID + 1)
    
    strGrade = Mid(pCommand, 1, cSIZE_GRADE)
    pCommand = Mid(pCommand, cSIZE_GRADE + 1)
    
    strLossCode = Mid(pCommand, 1)
    ReDim typDEFECT_DATA(intDefect_Count + 6)
    
    'For CAAPI change Grade-----Lucas20111127
            strGrade = PreJudgeGradeChange1(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = PreJudgeGradeChange2(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA, intDefect_Count)
            strGrade = PreJudgeGradeChange3(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), strPoint_Defect_Rank)
            strGrade = PostJudgeOtherRule1(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = PostJudgeOtherRule2(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = PostJudgeOtherRule3(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = PostJudgeGradeChange1(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = PostJudgeGradeChange2(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = RepairPointTimes(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = CheckPanelIDChangeGrade(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = ChangeGrade(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = ChangeGradeByDefectCode(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strGrade = FlagChangeGrade(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), pubJOB_INFO)
            strGrade = SKChange(strGrade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
'Lucas Ver.0.9.29 2012.06.01=================================For Light on Grade display
'    frmMain.flxMES_Data.TextMatrix(18, 1) = strGrade
    frmMain.lblPost_Judge.Caption = strGrade
'Lucas Ver.0.9.29 2012.06.01=================================For Light on Grade display
'    frmMain.flxMES_Data.TextMatrix(19, 1) = strLossCode
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_GRADE='" & strGrade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & strLossCode & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"
        
        dbMyDB.Execute (strQuery)
        
        dbMyDB.Close
    End If
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        .TextMatrix(intRow, 3) = strGrade
        .TextMatrix(intRow, 4) = strLossCode
        .TextMatrix(intRow, 5) = Get_Defect_Name(strLossCode)
        .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
    End With

    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortID, strStatus)
    
    If intPortID > 0 Then
        If Len(EQP.Get_BackLight_Value) > 5 Then
            strBackLight_Value = Left(EQP.Get_BackLight_Value, 5)
        Else
            intSpace = cSIZE_BACKLIGHT_VALUE - Len(EQP.Get_BackLight_Value)
            strBackLight_Value = EQP.Get_BackLight_Value & Space(intSpace)
        End If
        If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
            Call QUEUE.Put_Send_Command(intPortID, "QJPG" & strPanelID & strGrade & strLossCode & strBackLight_Value)
        Else
            strLossCode = Left(strLossCode, cSIZE_LOSSCODE)
            Call QUEUE.Put_Send_Command(intPortID, "QJPG" & strPanelID & strGrade & strLossCode)
        End If
    End If
    
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
'    Lucas.2012.03.08 Ver.0.9.13===========================After QSPO/QJPG sending ,Call make defect file
            If (frmMain.flxMES_Data.TextMatrix(18, 1) <> "GA") And (frmMain.flxMES_Data.TextMatrix(18, 1) <> "  ") Then
                Call Make_Defect_File
            End If
'    ===========================================End if
            Call EQP.Set_DEFECT_UPLOAD(True)
        End If
    Decode_Panel_Grade = "YAPG"
    
End Function

Private Sub Decode_EQ_State_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strRunningStatus        As String
    Dim strErrorText            As String
    
    strRunningStatus = Mid(pCommand, 5)
    Select Case strRunningStatus
    Case "RNC":
        strErrorText = "Not Contact"
    Case "RCO":
        strErrorText = "Contact"
    Case "REE":
        strErrorText = "Emergency Error"
    Case "RSE":
        strErrorText = "Safty Error"
    Case "RVC":
        strErrorText = "Vacuum Error"
    Case "RAE":
        strErrorText = "Auto alignment Error"
    Case "RRE":
        strErrorText = "Auto Re-alignment Error"
    Case "ROE":
        strErrorText = "Other Error"
    End Select
    
    frmMain.flxAPI_Information.TextMatrix(4, 1) = strRunningStatus & " : " & strErrorText
    
End Sub

Private Function Decode_EQ_State_Report(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim strRunningStatus        As String
    Dim strErrorText            As String
    
    strRunningStatus = Mid(pCommand, 5)
    Select Case strRunningStatus
    Case "RNC":
        strErrorText = "Not Contact"
    Case "RCO":
        strErrorText = "Contact"
    Case "REE":
        strErrorText = "Emergency Error"
    Case "RSE":
        strErrorText = "Safty Error"
    Case "RVC":
        strErrorText = "Vacuum Error"
    Case "RAE":
        strErrorText = "Auto alignment Error"
    Case "RRE":
        strErrorText = "Auto Re-alignment Error"
    Case "ROE":
        strErrorText = "Other Error"
    End Select
    
    frmMain.flxAPI_Information.TextMatrix(4, 1) = strRunningStatus & " : " & strErrorText
    
    Decode_EQ_State_Report = "YEQS"
    
End Function

Private Function Decode_Panel_Defect_Data(ByVal pPortID As Integer, ByVal pCommand As String) As String
    
    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
    Dim typPATTERN_LIST                 As PATTERN_LIST_DATA
    
    Dim strDefect_Code(1 To 3)          As String
    Dim strDATA_ADDRESS(1 To 3)         As String
    Dim strGATE_ADDRESS(1 To 3)         As String
        
    Dim strData                         As String
    
    Dim intDefect_Count                 As Integer
    Dim intIndex                        As Integer
    Dim intCol                          As Integer
    Dim intRow                          As Integer
    Dim intGrade_Count                  As Integer
    
    pCommand = Mid(pCommand, 5)
    
    If pCommand <> "" Then
        intDefect_Count = 0
        strData = pCommand
        While (pCommand <> "") And (intDefect_Count < 3)
            intDefect_Count = intDefect_Count + 1
            strData = Left(pCommand, 15)
            pCommand = Mid(pCommand, 16)
            
            strDefect_Code(intDefect_Count) = Left(strData, 5)
            strDATA_ADDRESS(intDefect_Count) = Mid(strData, 6, 5)
            strGATE_ADDRESS(intDefect_Count) = Right(strData, 5)
        Wend
        
        Call RANK_OBJ.Set_DEFECT_DATA_COUNT(intDefect_Count)
        For intIndex = 1 To intDefect_Count
            typRANK_DATA = Get_DEFECT_DATA_by_CODE(strDefect_Code(intIndex))
            With typRANK_DATA
                If Mid(typRANK_DATA.DEFECT_CODE, 2, 1) = "D" Then
                    If typRANK_DATA.DETAIL_DIVISION = "B" Then
                        Call RANK_OBJ.Add_TB_Count(typRANK_DATA.ACCUMULATION)
                    ElseIf typRANK_DATA.DETAIL_DIVISION = "D" Then
                        Call RANK_OBJ.Add_TD_Count(typRANK_DATA.ACCUMULATION)
                    End If
                End If
                Call RANK_OBJ.Set_DEFECT_DATA(intIndex, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, strDATA_ADDRESS, strGATE_ADDRESS, _
                                              "", 0, .ACCUMULATION)
            End With
        Next intIndex
    End If
    
    Decode_Panel_Defect_Data = "YPDD"
    
End Function
Private Sub Receive_Judge_Panel_Grade_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strCommand                  As String
    
    Dim intLength                   As Integer
    
    strCommand = "QSPO"
    
    intLength = cSIZE_PANELID - Len(Trim(frmMain.flxAlign_PanelID.TextMatrix(1, 0)))
    strCommand = strCommand & Trim(frmMain.flxAlign_PanelID.TextMatrix(1, 0)) & Space(intLength)
    
    Call QUEUE.Put_Send_Command(pPortID, strCommand)
    
End Sub
