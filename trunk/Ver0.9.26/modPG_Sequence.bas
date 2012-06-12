Attribute VB_Name = "modPG_Sequence"
Option Explicit

Public Sub PG_Sequence(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strReply                    As String
    Dim strDevice_State             As String
    
    Dim intPortID                   As Integer
    
    Call ENV.Get_Device_Data_by_Name("PG", intPortID, strDevice_State)
    
    If (strDevice_State = cDEVICE_ONLINE) Or (UCase(Left(pCommand, 4)) = "PDRF") Then
        Select Case UCase(Left(pCommand, 4))
        Case "PDRF":
            strReply = Decode_Online_PG(pPortID, pCommand)
        Case "RONG":                'On line report from PG
            strReply = Decode_Online_PG(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "ROFG":                'Off line report from PG
            strReply = Decode_Offline_PG(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "PSMG", "PSMY":                'Setting modify command reply
            Call Decode_Setting_Modify(pPortID, pCommand)
        Case "RBAM":                'Buzz and Message from PG
            strReply = Decode_Buzz_Message(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, strReply)
        Case "PPPO":                'PG Power ON reply
            Call Decode_Power_On(pPortID, pCommand)
        Case "PPPF":                'PG Power OFF reply
            Call Decode_Power_Off(pPortID, pCommand)
        Case "PPCC":                'Pattern change command reply
            Call Decode_Pattern_Change(pPortID, pCommand)
            frmJudge.picCurrent_Pattern.Enabled = True
        Case "RFDD":                'Coordinate command
            strReply = Decode_Coordinate(pPortID, pCommand)
        End Select
    End If
    
End Sub

Private Function Decode_Online_PG(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim strPanelDriveType               As String
    Dim strPFCD                         As String
    Dim strRS                           As String
    
    strRS = Mid(pCommand, 5, 1)
    
    If strRS = "0" Then
        Call State_Change(pPortID, "PG", cDEVICE_ONLINE)
        If frmMain.flxEQ_Information.TextMatrix(5, 1) <> "" Then
            Call ENV.Set_PG_Name("PG" & Right(frmMain.flxEQ_Information.TextMatrix(5, 1), 3))
        End If
    Else
    End If
    
'    If Get_Device_State("PG") = cDEVICE_OFFLINE Then
'        strPanelDriveType = Mid(pCommand, 5, 2)
'        strPFCD = Mid(pCommand, 7)
'        Call EQP.Set_PG_Panel_Drive_Type(strPanelDriveType)
'
'        Call State_Change(pPortID, "PG", cDEVICE_ONLINE)
'        Decode_Online_PG = "YONG"
'
'        If frmMain.flxEQ_Information.TextMatrix(2, 1) <> "" Then
'            If strPFCD <> Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5) Then
'                Call QUEUE.Put_Send_Command(pPortID, "QSMY" & Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5))
''                Call Show_Message("Data mismatch", "PFCD is not matched.")
'            End If
'        End If
'    End If
    
End Function

Private Function Decode_Offline_PG(ByVal pPortID As Integer, ByVal pCommand As String) As String

    If Get_Device_State("PG") = cDEVICE_ONLINE Then
        Call State_Change(pPortID, "PG", cDEVICE_OFFLINE)
        Decode_Offline_PG = "YOFG"
    End If
    
End Function

Private Sub Decode_Setting_Modify(ByVal pPortID As Integer, ByVal pCommand)

    Dim strPanelDriveType               As String
    Dim strPFCD                         As String
    Dim strRS                           As String
    Dim strStatus                       As String
    
    Dim intPortNo                       As Integer
    
    strPanelDriveType = Mid(pCommand, 5, 2)
    strRS = Right(pCommand, 1)
    
    Call EQP.Set_PG_Panel_Drive_Type(strPanelDriveType)
        
    If strRS = "1" Then
        Call Show_Message("PG Error", "PG Recipe change fail.")
        Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0010PG recipe change fail.")
        End If
'    Else
'        Call Power_On_PG
    End If
    
End Sub

Private Function Decode_Buzz_Message(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim strAlarmCode                    As String
    Dim strAlarmMessage                 As String
    Dim strStatus                       As String
    
    Dim intPortNo                       As Integer
    
    Decode_Buzz_Message = "YBAM"
    
    Call Show_Message("PG Alarm occurred.", "Received Message : " & Mid(pCommand, 5))
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QBAM0" & Mid(pCommand, 5))
    End If
    
End Function

Private Sub Decode_Power_On(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typPATTERN_DATA                 As PATTERN_LIST_DATA
    
    Dim strResponse                     As String
    Dim strPath                         As String
    Dim strFileName                     As String
    
    Dim intPTN_Index                    As Integer
    
    strResponse = Mid(pCommand, 5, 1)
    
    frmJudge.lblCurrent_PTN_Index.Caption = CInt(frmJudge.lblCurrent_PTN_Index.Caption) + 1
    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    With typPATTERN_DATA
        Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        Call Insert_Pattern_Start(RANK_OBJ.Get_Select_DEFECTCODE, typPATTERN_DATA.PATTERN_NAME)
    End With
    If typPATTERN_DATA.PATTERN_NAME <> "" Then
        strPath = App.PATH & "\Env\Standard_Info\"
        strFileName = typPATTERN_DATA.PATTERN_NAME & ".jpg"
        If Dir(strPath & strFileName) <> "" Then
            frmJudge.imgPG_Image.Picture = LoadPicture(strPath & strFileName)
        End If
    End If
    
    If (frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3) <> "0") And (frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3) <> "") Then
        frmJudge.tmrPattern_Delay.Interval = CLng(frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3))
        frmJudge.picCurrent_Pattern.Enabled = False
        frmJudge.tmrPattern_Delay.Enabled = True
    Else
        frmJudge.picCurrent_Pattern.Enabled = True
        frmJudge.tmrPattern_Delay.Enabled = False
    End If
    
    frmJudge.cmdGrade.Enabled = False
    Call EQP.Set_PATTERN_START_by_Index(intPTN_Index)
    
End Sub

Private Sub Decode_Power_Off(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typPATTERN_DATA                 As PATTERN_LIST_DATA
    
    Dim intPTN_Index                    As Integer
    
    Dim strResponse                     As String
    
    strResponse = Mid(pCommand, 5, 1)
    
    With typPATTERN_DATA
        intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
        Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        Call Insert_Pattern_End(RANK_OBJ.Get_Select_DEFECTCODE, .PATTERN_NAME)
    End With
    
End Sub

Private Sub Decode_Pattern_Change(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typPATTERN_DATA                 As PATTERN_LIST_DATA
    
    Dim strResponse                     As String
    Dim strPath                         As String
    Dim strFileName                     As String
    
    Dim intPTN_Index                    As Integer
    
    strResponse = Mid(pCommand, 5, 1)
    
    With typPATTERN_DATA
        intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
        Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        Call Insert_Pattern_End(RANK_OBJ.Get_Select_DEFECTCODE, .PATTERN_NAME)
    End With
    
    frmJudge.lblCurrent_PTN_Index.Caption = CInt(frmJudge.lblCurrent_PTN_Index.Caption) + 1
    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    With typPATTERN_DATA
        Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        Call Insert_Pattern_Start(RANK_OBJ.Get_Select_DEFECTCODE, .PATTERN_NAME)
    End With
    If typPATTERN_DATA.PATTERN_NAME <> "" Then
        strPath = App.PATH & "\Env\Standard_Info\"
        strFileName = typPATTERN_DATA.PATTERN_NAME & ".jpg"
        If Dir(strPath & strFileName) <> "" Then
            frmJudge.imgPG_Image.Picture = LoadPicture(strPath & strFileName)
        End If
    End If
    
    If (frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3) <> "0") And (frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3) <> "") Then
        frmJudge.tmrPattern_Delay.Interval = CLng(frmJudge.flxPG_Data.TextMatrix(intPTN_Index, 3))
        frmJudge.picCurrent_Pattern.Enabled = False
        frmJudge.tmrPattern_Delay.Enabled = True
    Else
        frmJudge.picCurrent_Pattern.Enabled = True
        frmJudge.tmrPattern_Delay.Enabled = False
    End If
        
    Call EQP.Set_PATTERN_START_by_Index(intPTN_Index)
    
End Sub

Private Function Decode_Coordinate(ByVal pPortID As Integer, ByVal pCommand As String) As String

    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
    Dim typPATTERN_LIST                 As PATTERN_LIST_DATA
    
    Dim strDATA_ADDRESS(1 To 3)         As String
    Dim strGATE_ADDRESS(1 To 3)         As String
        
    Dim strDefect_Code                  As String
    Dim strResponse                     As String
    
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
                    If (Trim(typRANK_DATA.Rank(intRankLevel)) <> "0") And (Trim(typRANK_DATA.Rank(intRankLevel)) <> "-") Then
                        .lblGrade(intRankLevel).Caption = RankLevel(intRankLevel)
                        .optSpec_Value(intRankLevel).Caption = typRANK_DATA.Rank(intRankLevel)
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
    
    strResponse = "0"
    Decode_Coordinate = "YFDD" & strResponse
    
End Function
