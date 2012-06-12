Attribute VB_Name = "modGrade_Calculate"
Option Explicit

Public Function Check_TFT_CF_PanelID(ByVal pPanelID As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strAlarm_Msg                As String
    
    Dim intIndex                    As Integer
    
    Dim bolEnable                   As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    bolEnable = False
    For intIndex = 1 To 10
        If EQP.Get_Control_Data("checkTFTPanelID" & intIndex) = "E" Then
            bolEnable = True
        End If
    Next intIndex
    
    If bolEnable = True Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strAlarm_Msg = ""
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM ABNORMAL_PANEL WHERE "
            strQuery = strQuery & "PANELID ='" & pPanelID & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strAlarm_Msg = lstRecord.Fields("ALARM_TEXT")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strAlarm_Msg = ""
    End If
    
    Check_TFT_CF_PanelID = strAlarm_Msg
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Check_TFT_CF_PanelID", ErrMsg)
    
End Function

Public Function Check_MES_Data(pCST_DATA As CST_INFO_ELEMENTS, pPANEL_DATA As PANEL_INFO_ELEMENTS, pJOB_DATA As JOB_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim arrFind_Data(1 To 50)       As Boolean
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strAlarm_Msg                As String
    
    Dim bolFind_False               As Boolean
    
    Dim intIndex                    As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("CheckMesData") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM ABNORMAL_MES_DATA WHERE "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_DATA.PROCESS_NUM & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strAlarm_Msg = ""
                While lstRecord.EOF = False
                    If strAlarm_Msg = "" Then
                        With lstRecord
                            For intIndex = 1 To 50
                                arrFind_Data(intIndex) = True
                            Next intIndex
                            If (Trim(.Fields("PROCESSNUM")) = "") Or (Trim(.Fields("PROCESSNUM")) = pCST_DATA.PROCESS_NUM) Then
                                arrFind_Data(1) = True
                            Else
                                arrFind_Data(1) = False
                            End If
                            
                            If Trim(.Fields("DEST_FAB")) = "" Or Trim(pCST_DATA.DESTINATION_FAB) = Trim(.Fields("DEST_FAB")) Then
                                
                               arrFind_Data(2) = True
                               
                            Else
                                arrFind_Data(2) = False
                            End If
                            
                            If Trim(.Fields("RMANO")) = "" Or Trim(pCST_DATA.RMANO) = Trim(.Fields("RMANO")) Then
                                arrFind_Data(3) = True
                            Else
                                arrFind_Data(3) = False
                            End If
                            
                            If Trim(.Fields("OQCNO")) = "" Or Trim(pCST_DATA.OQCNO) = Trim(.Fields("OQCNO")) Then
                                arrFind_Data(4) = True
                            Else
                                arrFind_Data(4) = False
                            End If
                            
                            If Trim(.Fields("PANELID")) = "" Or Trim(pPANEL_DATA.PANELID) = Trim(.Fields("PANELID")) Then
                                arrFind_Data(5) = True
                            Else
                                arrFind_Data(5) = False
                            End If
                            
                            If Trim(.Fields("LIGHT_ON_PANEL_GRADE")) = "" Or Trim(pPANEL_DATA.LIGHT_ON_PANEL_GRADE) = Trim(.Fields("LIGHT_ON_PANEL_GRADE")) Then
                                arrFind_Data(6) = True
                            Else
                                arrFind_Data(6) = False
                            End If
                            
                            If Trim(.Fields("LIGHT_ON_REASON_CODE")) = "" Or Trim(pPANEL_DATA.LIGHT_ON_REASON_CODE) = Trim(.Fields("LIGHT_ON_REASON_CODE")) Then
                                arrFind_Data(7) = True
                            Else
                                arrFind_Data(7) = False
                            End If
                            
                            If Trim(.Fields("CELL_LINE_RESCUE_FLAG")) = "" Or Trim(pPANEL_DATA.CELL_LINE_RESCUE_FLAG) = Trim(.Fields("CELL_LINE_RESCUE_FLAG")) Then
                                arrFind_Data(8) = True
                            Else
                                arrFind_Data(8) = False
                            End If
                            
                            If Trim(.Fields("CELL_REPAIR_JUDGE_GRADE")) = "" Or Trim(pPANEL_DATA.CELL_REPAIR_JUDGE_GRADE) = Trim(.Fields("CELL_REPAIR_JUDGE_GRADE")) Then
                                
                                arrFind_Data(9) = True
                           
                            Else
                                arrFind_Data(9) = False
                            End If
                            
                            If Trim(.Fields("TFT_REPAIR_GRADE")) = "" Or Trim(pPANEL_DATA.TFT_REPAIR_GRADE) = Trim(.Fields("TFT_REPAIR_GRADE")) Then
                               
                                arrFind_Data(10) = True
                            Else
                                arrFind_Data(10) = False
                            End If
                            
                            If Trim(.Fields("CF_PANELID")) = "" Or Trim(pPANEL_DATA.CF_PANELID) = Trim(.Fields("CF_PANELID")) Then
                                arrFind_Data(11) = True
                            Else
                                arrFind_Data(11) = False
                            End If
                            
                            If Trim(.Fields("CF_PANEL_OX_INFORMATION")) = "" Or Trim(pPANEL_DATA.CF_PANEL_OX_INFORMATION) = Trim(.Fields("CF_PANEL_OX_INFORMATION")) Then
                               arrFind_Data(12) = True
                            Else
                                arrFind_Data(12) = False
                            End If
                            
                            If Trim(.Fields("PANEL_OWNER_TYPE")) = "" Or Trim(pPANEL_DATA.PANEL_OWNER_TYPE) = Trim(.Fields("PANEL_OWNER_TYPE")) Then
                             
                                    arrFind_Data(13) = True
                       
                            Else
                                arrFind_Data(13) = False
                            End If
                            
                            If Trim(.Fields("ABNORMAL_CF")) = "" Or Trim(pPANEL_DATA.ABNORMAL_CF) = Trim(.Fields("ABNORMAL_CF")) Then
                             
                                    arrFind_Data(14) = True
                            
                            Else
                                arrFind_Data(14) = False
                            End If
                            
                            If Trim(.Fields("ABNORMAL_TFT")) = "" Or Trim(pPANEL_DATA.ABNORMAL_TFT) = Trim(.Fields("ABNORMAL_TFT")) Then
                               
                                    arrFind_Data(15) = True
                            Else
                                arrFind_Data(15) = False
                            End If
                            
                            If Trim(.Fields("ABNORMAL_LCD")) = "" Or Trim(pPANEL_DATA.ABNORMAL_LCD) = Trim(.Fields("ABNORMAL_LCD")) Then
                                
                                arrFind_Data(16) = True
                            Else
                                arrFind_Data(16) = False
                            End If
                            
                            If Trim(.Fields("GROUPID")) = "" Or Trim(pPANEL_DATA.GROUP_ID) = Trim(.Fields("GROUPID")) Then
                                
                                    arrFind_Data(17) = True
                              
                            Else
                                arrFind_Data(17) = False
                            End If
                            
                            If Trim(.Fields("REPAIR_REWORK_COUNT")) = "" Or Trim(pPANEL_DATA.REPAIR_REWORK_COUNT) = Trim(.Fields("REPAIR_REWORK_COUNT")) Then
                               
                                    arrFind_Data(18) = True
                              
                            Else
                                arrFind_Data(18) = False
                            End If
                            
                            If Trim(.Fields("POLARIZER_REWORK_COUNT")) = "" Or Trim(pPANEL_DATA.POLARIZER_REWORK_COUNT) = Trim(.Fields("POLARIZER_REWORK_COUNT")) Then
                               
                                    arrFind_Data(19) = True
                              
                            Else
                                arrFind_Data(19) = False
                            End If
                            
                            If Trim(.Fields("X_TOTAL_PIXEL")) = "" Or Trim(pPANEL_DATA.X_TOTAL_PIXEL) = Trim(.Fields("X_TOTAL_PIXEL")) Then
                                
                                    arrFind_Data(20) = True
                            Else
                                arrFind_Data(20) = False
                            End If
                            
                            If Trim(.Fields("Y_TOTAL_PIXEL")) = "" Or Trim(pPANEL_DATA.Y_TOTAL_PIXEL) = Trim(.Fields("Y_TOTAL_PIXEL")) Then
                               
                                    arrFind_Data(21) = True
                           
                            Else
                                arrFind_Data(21) = False
                            End If
                            
                            If Trim(.Fields("LCD_Q_TAB_LOT_GROUPID")) = "" Or Trim(pPANEL_DATA.LCD_Q_TAP_LOT_GROUPID) = Trim(.Fields("LCD_Q_TAB_LOT_GROUPID")) Then
                               
                                    arrFind_Data(22) = True
                           
                            Else
                                arrFind_Data(22) = False
                            End If
                            
                            If Trim(.Fields("SK_FLAG")) = "" Or Trim(pPANEL_DATA.SK_FLAG) = Trim(.Fields("SK_FLAG")) Then
                               
                                    arrFind_Data(23) = True
                             
                            Else
                                arrFind_Data(23) = False
                            End If
                            
                            If Trim(.Fields("CF_R_DEFECT_CODE")) = "" Or Trim(pPANEL_DATA.CF_R_DEFECT_CODE) = Trim(.Fields("CF_R_DEFECT_CODE")) Then
                                
                                arrFind_Data(24) = True
                            
                            Else
                                arrFind_Data(24) = False
                            End If
                            
                            If Trim(.Fields("ODK_AK_FLAG")) = "" Or Trim(pPANEL_DATA.ODK_AK_FLAG) = Trim(.Fields("ODK_AK_FLAG")) Then
                               
                                    arrFind_Data(25) = True
                            Else
                                arrFind_Data(25) = False
                            End If
                            
                            If Trim(.Fields("BPAM_REWORK_FLAG")) = "" Or Trim(pPANEL_DATA.BPAM_REWORK_FLAG) = Trim(.Fields("BPAM_REWORK_FLAG")) Then
                               
                                    arrFind_Data(26) = True
                                
                            Else
                                arrFind_Data(26) = False
                            End If
                            
                            If Trim(.Fields("LCD_BRIGHT_DOT_FLAG")) = "" Or Trim(pPANEL_DATA.LCD_BRIGHT_DOT_FLAG) = Trim(.Fields("LCD_BRIGHT_DOT_FLAG")) Then
                              
                                    arrFind_Data(27) = True
                              
                            Else
                                arrFind_Data(27) = False
                            End If
                            
                            If Trim(.Fields("CF_PS_HEIGHT_ERR_FLAG")) = "" Or Trim(pPANEL_DATA.CF_PANEL_OX_INFORMATION) = Trim(.Fields("CF_PS_HEIGHT_ERR_FLAG")) Then
                                
                                    arrFind_Data(28) = True
                              
                            Else
                                arrFind_Data(28) = False
                            End If
                            
                            If Trim(.Fields("PI_INSPECTION_NG_FLAG")) = "" Or Trim(pPANEL_DATA.PI_INSPECTION_NG_FLAG) = Trim(.Fields("PI_INSPECTION_NG_FLAG")) Then
                                
                                    arrFind_Data(29) = True
                                
                            Else
                                arrFind_Data(29) = False
                            End If
                            
                            If Trim(.Fields("PI_OVER_BAKE_FLAG")) = "" Or Trim(pPANEL_DATA.PI_OVER_BAKE_FLAG) = Trim(.Fields("PI_OVER_BAKE_FLAG")) Then
                                
                               arrFind_Data(30) = True
                           
                            Else
                                arrFind_Data(30) = False
                            End If
                            
                            If Trim(.Fields("PI_OVER_Q_TIME_FLAG")) = "" Or Trim(pPANEL_DATA.PI_OVER_Q_TIME_FLAG) = Trim(.Fields("PI_OVER_Q_TIME_FLAG")) Then
                                
                                    arrFind_Data(31) = True
                           
                            Else
                                arrFind_Data(31) = False
                            End If
                            
                            If Trim(.Fields("ODF_OVER_BAKE_FLAG")) = "" Or Trim(pPANEL_DATA.ODF_OVER_BAKE_FLAG) = Trim(.Fields("ODF_OVER_BAKE_FLAG")) Then
                               
                                    arrFind_Data(32) = True
                            
                            Else
                                arrFind_Data(32) = False
                            End If
                            
                            If Trim(.Fields("ODF_OVER_Q_TIME_FLAG")) = "" Or Trim(pPANEL_DATA.ODF_OVER_Q_TIME_FLAG) = Trim(.Fields("ODF_OVER_Q_TIME_FLAG")) Then
                                
                                    arrFind_Data(33) = True
                            
                            Else
                                arrFind_Data(33) = False
                            End If
                            
                            If Trim(.Fields("HVA_OVER_BAKE_FLAG")) = "" Or Trim(pPANEL_DATA.HVA_OVER_BAKE_FLAG) = Trim(.Fields("HVA_OVER_BAKE_FLAG")) Then
                                
                                    arrFind_Data(34) = True
                            Else
                                arrFind_Data(34) = False
                            End If
                            
                            If Trim(.Fields("HVA_OVER_Q_TIME_FLAG")) = "" Or Trim(pPANEL_DATA.HVA_OVER_Q_TIME_FLAG) = Trim(.Fields("HVA_OVER_Q_TIME_FLAG")) Then
                               
                                    arrFind_Data(35) = True
                              
                            Else
                                arrFind_Data(35) = False
                            End If
                            
                            If Trim(.Fields("SEAL_INSPECTION_FLAG")) = "" Or Trim(pPANEL_DATA.SEAL_INSPECTION_FLAG) = Trim(.Fields("SEAL_INSPECTION_FLAG")) Then
                               
                                arrFind_Data(36) = True
                            Else
                                arrFind_Data(36) = False
                            End If
                            
                            If Trim(.Fields("ODF_CHECKER_FLAG")) = "" Or Trim(pPANEL_DATA.ODF_CHECKER_FLAG) = Trim(.Fields("ODF_CHECKER_FLAG")) Then
                                
                               arrFind_Data(37) = True
                            
                            Else
                                arrFind_Data(37) = False
                            End If
                            
                            If Trim(.Fields("ODF_DOOR_OPEN_FLAG")) = "" Or Trim(pPANEL_DATA.ODF_DOOR_OPEN_FLAG) = Trim(.Fields("ODF_DOOR_OPEN_FLAG")) Then
                               
                                    arrFind_Data(38) = True
                            Else
                                arrFind_Data(38) = False
                            End If
                            
                            If Trim(.Fields("JOB_JUDGE")) = "" Or Trim(pJOB_DATA.JOB_JUDGE) = Trim(.Fields("JOB_JUDGE")) Then
                               
                                    arrFind_Data(39) = True
                            
                            Else
                                arrFind_Data(39) = False
                            End If
                            
                            If Trim(.Fields("JOB_GRADE")) = "" Or Trim(pJOB_DATA.JOB_GRADE) = Trim(.Fields("JOB_GRADE")) Then
                               
                               arrFind_Data(40) = True
                            Else
                                arrFind_Data(40) = False
                            End If
                            
                            arrFind_Data(41) = True
                            arrFind_Data(42) = True
                            arrFind_Data(43) = True
                            arrFind_Data(44) = True
                            arrFind_Data(45) = True
                            
                            If Trim(.Fields("SAMPLING_SLOT_FLAG")) = "" Or Trim(pJOB_DATA.SAMPLING_SLOT_FLAG) = Trim(.Fields("SAMPLING_SLOT_FLAG")) Then
                               
                                    arrFind_Data(46) = True
                            
                            Else
                                arrFind_Data(46) = False
                            End If
                            
                            If Trim(.Fields("NEED_GRINDING_FLAG")) = "" Or Trim(pJOB_DATA.NEED_GRINDING_FLAG) = Trim(.Fields("NEED_GRINDING_FLAG")) Then
                                
                                    arrFind_Data(47) = True
                              
                            Else
                                arrFind_Data(47) = False
                            End If
                            
                            If Trim(.Fields("SMALL_MULTI_PANEL_FLAG")) = "" Or Trim(pJOB_DATA.SMALL_MULTI_PANEL_FLAG) = Trim(.Fields("SMALL_MULTI_PANEL_FLAG")) Then
                                
                                    arrFind_Data(48) = True
                             
                            Else
                                arrFind_Data(48) = False
                            End If
                            
                            If Trim(.Fields("CST_SETTING_CODE")) = "" Or Trim(pJOB_DATA.CASSETTE_SETTING_CODE) = Trim(.Fields("CST_SETTING_CODE")) Then
                               
                                    arrFind_Data(49) = True
                            
                            Else
                                arrFind_Data(49) = False
                            End If
                            
                            If Trim(.Fields("ABNORMAL_FLAG_CODE")) = "" Or Trim(pJOB_DATA.ABNORMAL_FLAG_CODE) = Trim(.Fields("ABNORMAL_FLAG_CODE")) Then
                                
                                    arrFind_Data(50) = True
                            Else
                                arrFind_Data(50) = False
                            End If
                        End With
                        bolFind_False = False
                        For intIndex = 1 To 50
                            If arrFind_Data(intIndex) = False Then
                                bolFind_False = True
                            End If
                        Next intIndex
                        If bolFind_False = False Then
                            strAlarm_Msg = lstRecord.Fields("ALARM_TEXT")
                        End If
                    End If
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strAlarm_Msg = ""
    End If
    
    Check_MES_Data = strAlarm_Msg
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Check_MES_Data", ErrMsg)
    
End Function

Public Function Assign_Grade(pCST_DATA As CST_INFO_ELEMENTS, pPANEL_DATA As PANEL_INFO_ELEMENTS) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    
    Dim intPriority                 As Integer
    Dim intIndex                    As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strNew_Judge = ""
    
    For intIndex = 1 To 5
        If EQP.Get_Control_Data("AssignGrade" & intIndex) = "E" Then
            strDB_Path = App.PATH & "\DB\"
            strDB_FileName = "STANDARD_INFO.mdb"
'            strNew_Judge = ""
            intPriority = 100
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
                strQuery = "SELECT * FROM ASSIGN_GRADE WHERE "
                strQuery = strQuery & "PFCD = '" & Mid(pCST_DATA.PFCD, 3, 5) & "' AND "
                strQuery = strQuery & "PROCESSNUM = '" & pCST_DATA.PROCESS_NUM & "' AND "
                strQuery = strQuery & "PANELID = '" & pubPANEL_INFO.PANELID & "' AND "
                strQuery = strQuery & "PRIORITY = " & intIndex
                        
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    lstRecord.MoveFirst
                    If lstRecord.Fields("PanelID") = pubPANEL_INFO.PANELID Then
                    strNew_Judge = lstRecord.Fields("NEW_GRADE")
                    Else
                    strNew_Judge = ""
                    End If
                    lstRecord.MoveNext
                End If
                lstRecord.Close
                
                dbMyDB.Close
            End If
        End If
    Next intIndex
    
    Assign_Grade = strNew_Judge
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Assigh_Grade", ErrMsg)
    
End Function

Public Sub Get_Rank_Data(ByVal pPROCESSNUM As String, pRANK_DATA As RANK_DATA_STRUCTURE, pGRADE_DATA() As GRADE_DATA_STRUCTURE, ByVal pDEFECT_CODE As String, pData_Count As Integer)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset

    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intRecord_Count             As Integer
    Dim intRecord_Index             As Integer
    
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
'        Select Case pPROCESSNUM
'        Case "3000":
'            strQuery = strQuery & "RANK_DIVISION = 'LOI-1' AND "
'        Case "3650":
'            strQuery = strQuery & "RANK_DIVISION = 'RLOI-1' AND "
'        Case "4600":
'            strQuery = strQuery & "RANK_DIVISION = 'LOI-2' AND "
'        Case "4650":
'            strQuery = strQuery & "RANK_DIVISION = 'RLOI-2' AND "
'        End Select
        strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_CODE & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With pRANK_DATA
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
                '============Leo 2012.05.22 Add Rank Level End
                .ODF = lstRecord.Fields("ODF")
                .PRIORITY = lstRecord.Fields("PRIORITY")
                .POP_UP = lstRecord.Fields("POP_UP")
            End With
        End If
        lstRecord.Close
        
       strQuery = "SELECT * FROM GRADE_DATA WHERE "
       strQuery = strQuery & "RANK_DIVISION = '" & pRANK_DATA.RANK_DIVISION & "' AND "
       strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_CODE & "'"
       
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            intRecord_Count = lstRecord.RecordCount
            ReDim pGRADE_DATA(intRecord_Count)
            pData_Count = intRecord_Count
            intRecord_Index = 0
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRecord_Index = intRecord_Index + 1
                With pGRADE_DATA(intRecord_Index)
                    .RANK_DIVISION = lstRecord.Fields("RANK_DIVISION")
                    .DEFECT_CODE = lstRecord.Fields("DEFECT_CODE")
                    .RANK = lstRecord.Fields("RANK")
                    .GRADE = lstRecord.Fields("GRADE")
                End With
                
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Sub

Public Sub ACCUMULATE(pCST_DATA As CST_INFO_ELEMENTS, pDEFECT_DATA As DEFECT_DATA_STRUCTURE, ByVal pRow As Integer)
    
    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim typRANK_DATA                As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()             As GRADE_DATA_STRUCTURE
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strRank                     As String
    Dim strGrade                    As String
    Dim strQuery                    As String
    
    Dim intAccumulation             As Integer
    Dim intGrade_Count              As Integer
    Dim intIndex                    As Integer
    Dim intPoint_Defect_Total       As Integer
    
    Call Get_Rank_Data(pCST_DATA.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, pDEFECT_DATA.DEFECT_CODE, intGrade_Count)
    pDEFECT_DATA.DETAIL_DIVISION = typRANK_DATA.DETAIL_DIVISION
    pDEFECT_DATA.PRIORITY = typRANK_DATA.PRIORITY
    
    If typRANK_DATA.DEFECT_TYPE = "P" Then
        If typRANK_DATA.USE_XY <> "X" Then
            'Accumulation
            intAccumulation = CInt(typRANK_DATA.ACCUMULATION)
            frmJudge.flxDefect_List.TextMatrix(pRow, 10) = intAccumulation
            pDEFECT_DATA.ACCUMULATION = intAccumulation
            Call Add_Point_Defect_Total(pDEFECT_DATA, intAccumulation)
            intPoint_Defect_Total = Get_Point_Defect_Total(pDEFECT_DATA.DEFECT_CODE, pDEFECT_DATA.PANELID)
            
            If typRANK_DATA.DETAIL_DIVISION = "B" Then
                Call RANK_OBJ.Add_TB_Count(intAccumulation)
            ElseIf typRANK_DATA.DETAIL_DIVISION = "D" Then
                Call RANK_OBJ.Add_TD_Count(intAccumulation)
            End If
            
            If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, intPoint_Defect_Total)
                pDEFECT_DATA.RANK = strRank
                pDEFECT_DATA.GRADE = strGrade
                Call RANK_OBJ.Set_DEFECT_RANK(pDEFECT_DATA.DEFECT_CODE, strRank, pDEFECT_DATA.DATA_ADDRESS, pDEFECT_DATA.GATE_ADDRESS)
                Call RANK_OBJ.Set_DEFECT_GRADE(pDEFECT_DATA.DEFECT_CODE, pDEFECT_DATA.DATA_ADDRESS, pDEFECT_DATA.GATE_ADDRESS, strGrade)
                pDEFECT_DATA.RANK = strRank
                If strGrade = "" Then
                    strGrade = frmMain.lblPre_Judge.Caption
                End If
                pDEFECT_DATA.GRADE = strGrade
                frmJudge.flxDefect_List.TextMatrix(pRow, 9) = strRank
                Call SaveLog("ACCUMULATE", pDEFECT_DATA.DEFECT_CODE & "'s RANK : " & strRank & ", GRADE : " & strGrade)
            End If
        Else
            Call Add_Other_Defect_Data(pDEFECT_DATA)
            'Not Accumulation
            If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, intPoint_Defect_Total)
                pDEFECT_DATA.RANK = strRank
                pDEFECT_DATA.GRADE = strGrade
                Call RANK_OBJ.Set_DEFECT_RANK(pDEFECT_DATA.DEFECT_CODE, strRank, pDEFECT_DATA.DATA_ADDRESS, pDEFECT_DATA.GATE_ADDRESS)
                Call RANK_OBJ.Set_DEFECT_GRADE(pDEFECT_DATA.DEFECT_CODE, pDEFECT_DATA.DATA_ADDRESS, pDEFECT_DATA.GATE_ADDRESS, strGrade)
                pDEFECT_DATA.RANK = strRank
                pDEFECT_DATA.GRADE = strGrade
                frmJudge.flxDefect_List.TextMatrix(pRow, 9) = strRank
                Call SaveLog("ACCUMULATE", pDEFECT_DATA.DEFECT_CODE & "'s RANK : " & strRank & ", GRADE : " & strGrade)
            End If
        End If
    Else
        Call Add_Other_Defect_Data(pDEFECT_DATA)
    End If

    Call Update_Defect_Grade(pDEFECT_DATA)
    
End Sub

Public Sub Get_Rank(pRANK_DATA As RANK_DATA_STRUCTURE, pGRADE_DATA() As GRADE_DATA_STRUCTURE, ByVal pArray_Count As Integer, _
                    pRank As String, pGrade As String, ByVal pDEFECT_TOTAL As Integer)

    Dim intIndex                As Integer
        '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

    pRank = ""
    pGrade = ""
    With pRANK_DATA
        If (.DEFECT_CODE <> "CDBDD") And _
           (.DEFECT_CODE <> "CDDKD") And _
           (.DEFECT_CODE <> "CDBDD") Then
            If (.DEFECT_CODE = "CDBTT") Or (.DEFECT_CODE = "CDDKT") Then
                If pDEFECT_TOTAL > 1 Then
'============Leo 2012.05.22 Add Rank Level Start
                    For intRankLevel = 0 To UBound(RankLevel)
                        If (pRank = "") And (IsNumeric(.Rank(intRankLevel)) = True) Then
                            If pDEFECT_TOTAL <= CInt(.Rank(intRankLevel)) Then
                                pRank = RankLevel(intRankLevel)
                            End If
                        End If
                    Next intRankLevel
                            

'                    If IsNumeric(.RANK_Y) = True Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_Y) Then
'                            pRank = "Y"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_L) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_L) Then
'                            pRank = "L"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_K) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_K) Then
'                            pRank = "K"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_C) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_C) Then
'                            pRank = "C"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_S) = True) Then
'                        If pDEFECT_TOTAL > CInt(.RANK_S) Then
'                            pRank = "S"
'                        End If
'                    End If
        '============Leo 2012.05.22 Add Rank Level End
                    For intIndex = 1 To pArray_Count
                        With pGRADE_DATA(intIndex)
                            If (pGrade = "") And (.DEFECT_CODE = pRANK_DATA.DEFECT_CODE) And (InStr(.RANK, pRank) > 0) Then
                                pGrade = .GRADE
                            End If
                        End With
                    Next intIndex
                End If
            Else
                If pDEFECT_TOTAL > 0 Then
                '============Leo 2012.05.22 Add Rank Level Start
                    For intRankLevel = 0 To UBound(RankLevel)
                        If (pRank = "") And (IsNumeric(.Rank(intRankLevel)) = True) Then
                            If pDEFECT_TOTAL <= CInt(.Rank(intRankLevel)) Then
                                pRank = RankLevel(intRankLevel)
                            End If
                        End If
                    Next intRankLevel
                    
'                    If IsNumeric(.RANK_Y) = True Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_Y) Then
'                            pRank = "Y"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_L) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_L) Then
'                            pRank = "L"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_K) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_K) Then
'                            pRank = "K"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_C) = True) Then
'                        If pDEFECT_TOTAL <= CInt(.RANK_C) Then
'                            pRank = "C"
'                        End If
'                    End If
'                    If (pRank = "") And (IsNumeric(.RANK_S) = True) Then
'                        If pDEFECT_TOTAL > CInt(.RANK_S) Then
'                            pRank = "S"
'                        End If
'                    End If
                 '============Leo 2012.05.22 Add Rank Level End
                    For intIndex = 1 To pArray_Count
                        With pGRADE_DATA(intIndex)
                            If (pGrade = "") And (.DEFECT_CODE = pRANK_DATA.DEFECT_CODE) And (InStr(.RANK, pRank) > 0) Then
                                pGrade = .GRADE
                            End If
                        End With
                    Next intIndex
                End If
            End If
        Else
            If pDEFECT_TOTAL > 1 Then
             '============Leo 2012.05.22 Add Rank Level Start
                    For intRankLevel = 0 To UBound(RankLevel)
                        If (pRank = "") And (IsNumeric(.Rank(UBound(RankLevel) - intRankLevel)) = True) Then
                            If pDEFECT_TOTAL <= CInt(.Rank(UBound(RankLevel) - intRankLevel)) Then
                                pRank = RankLevel(UBound(RankLevel) - intRankLevel)
                            End If
                        End If
                    Next intRankLevel
                    
'                If IsNumeric(.RANK_S) = True Then
'                    If pDEFECT_TOTAL <= CInt(.RANK_S) Then
'                        pRank = "S"
'                    End If
'                End If
'                If (pRank = "") And (IsNumeric(.RANK_C) = True) Then
'                    If pDEFECT_TOTAL <= CInt(.RANK_C) Then
'                        pRank = "C"
'                    End If
'                End If
'                If (pRank = "") And (IsNumeric(.RANK_K) = True) Then
'                    If pDEFECT_TOTAL <= CInt(.RANK_K) Then
'                        pRank = "K"
'                    End If
'                End If
'                If (pRank = "") And (IsNumeric(.RANK_L) = True) Then
'                    If pDEFECT_TOTAL <= CInt(.RANK_L) Then
'                        pRank = "L"
'                    End If
'                End If
'                If (pRank = "") And (IsNumeric(.RANK_Y) = True) Then
'                    If pDEFECT_TOTAL <= CInt(.RANK_Y) Then
'                        pRank = "Y"
'                    End If
'                End If
                '============Leo 2012.05.22 Add Rank Level End
                For intIndex = 1 To pArray_Count
                    With pGRADE_DATA(intIndex)
                        If (pGrade = "") And (.DEFECT_CODE = pRANK_DATA.DEFECT_CODE) And (InStr(.RANK, pRank) > 0) Then
                            pGrade = .GRADE
                        End If
                    End With
                Next intIndex
            End If
        End If
    End With
    
    
End Sub

Public Function Get_Grade_by_Rank(pGRADE_DATA() As GRADE_DATA_STRUCTURE, ByVal pArray_Count As Integer, ByVal pDEFECT_CODE As String, ByVal pRank As String) As String

    Dim intIndex                    As Integer
    
    For intIndex = 1 To pArray_Count
        With pGRADE_DATA(intIndex)
            If (.DEFECT_CODE = pDEFECT_CODE) And (InStr(.RANK, pRank) > 0) Then
                Get_Grade_by_Rank = .GRADE
            End If
        End With
    Next intIndex

End Function

Public Sub Add_Point_Defect_Total(pDEFECT_DATA As DEFECT_DATA_STRUCTURE, ByVal pDEFECT_COUNT As Integer)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strQuery                    As String
    Dim strCurrent_KEYID            As String
    
    Dim intPoint_Defect_Total       As Integer
    Dim intIndex                    As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    strCurrent_KEYID = RANK_OBJ.Get_Current_KEYID
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "PANELID = '" & pDEFECT_DATA.PANELID & "' AND "
        strQuery = strQuery & "KEYID = '" & strCurrent_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            strPath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        If strPath <> "" Then
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
        End If
        If (strPath <> "") And (strFileName <> "") Then
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strPath & strFileName)
                
                strQuery = "INSERT INTO DEFECT_DATA VALUES ("
                strQuery = strQuery & "'" & pDEFECT_DATA.PANELID & "', "
                strQuery = strQuery & pDEFECT_DATA.DEFECT_NO & ", "
                strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_CODE & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_NAME & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DETAIL_DIVISION & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.COLOR & "', "
                strQuery = strQuery & pDEFECT_DATA.GRAY_LEVEL & ", "
                For intIndex = 1 To 2
                    strQuery = strQuery & "'" & pDEFECT_DATA.GATE_ADDRESS(intIndex) & "', "
                    strQuery = strQuery & "'" & pDEFECT_DATA.DATA_ADDRESS(intIndex) & "', "
                Next intIndex
                strQuery = strQuery & "'" & pDEFECT_DATA.GATE_ADDRESS(3) & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DATA_ADDRESS(3) & "', "
                strQuery = strQuery & pDEFECT_COUNT & ", "
                strQuery = strQuery & "'" & pDEFECT_DATA.GRADE & "')"
                
                dbMyDB.Execute (strQuery)
                
                strQuery = "SELECT * FROM DEFECT_COUNT WHERE "
                strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_DATA.DEFECT_CODE & "'"
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    intPoint_Defect_Total = lstRecord.Fields("DEFECT_COUNT") + pDEFECT_COUNT
                    lstRecord.Close
                    
                    strQuery = "UPDATE DEFECT_COUNT SET "
                    strQuery = strQuery & "DEFECT_COUNT = " & intPoint_Defect_Total & " WHERE "
                    strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_DATA.DEFECT_CODE & "'"
                    
                    dbMyDB.Execute (strQuery)
                Else
                    lstRecord.Close
                    
                    strQuery = "INSERT INTO DEFECT_COUNT VALUES ("
                    strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_CODE & "', "
                    strQuery = strQuery & pDEFECT_COUNT & ")"
                    
                    dbMyDB.Execute (strQuery)
                End If
                
                dbMyDB.Close
            End If
        Else
        End If
    End If
    
End Sub

Public Sub Add_Other_Defect_Data(pDEFECT_DATA As DEFECT_DATA_STRUCTURE)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strQuery                    As String
    Dim strCurrent_KEYID            As String
    
    Dim intPoint_Defect_Total       As Integer
    Dim intIndex                    As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    strCurrent_KEYID = RANK_OBJ.Get_Current_KEYID
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "PANELID = '" & pDEFECT_DATA.PANELID & "' AND "
        strQuery = strQuery & "KEYID = '" & strCurrent_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            strPath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        If strPath <> "" Then
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
        End If
        If (strPath <> "") And (strFileName <> "") Then
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strPath & strFileName)
                
                strQuery = "INSERT INTO DEFECT_DATA VALUES ("
                strQuery = strQuery & "'" & pDEFECT_DATA.PANELID & "', "
                strQuery = strQuery & pDEFECT_DATA.DEFECT_NO & ", "
                strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_CODE & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_NAME & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DETAIL_DIVISION & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.COLOR & "', "
                strQuery = strQuery & pDEFECT_DATA.GRAY_LEVEL & ", "
                For intIndex = 1 To 2
                    strQuery = strQuery & "'" & pDEFECT_DATA.GATE_ADDRESS(intIndex) & "', "
                    strQuery = strQuery & "'" & pDEFECT_DATA.DATA_ADDRESS(intIndex) & "', "
                Next intIndex
                strQuery = strQuery & "'" & pDEFECT_DATA.GATE_ADDRESS(3) & "', "
                strQuery = strQuery & "'" & pDEFECT_DATA.DATA_ADDRESS(3) & "', "
                strQuery = strQuery & "1, "
                strQuery = strQuery & "'" & pDEFECT_DATA.GRADE & "')"
                
                dbMyDB.Execute (strQuery)
                
                strQuery = "SELECT * FROM DEFECT_COUNT WHERE "
                strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_DATA.DEFECT_CODE & "'"
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    intPoint_Defect_Total = lstRecord.Fields("DEFECT_COUNT") + 1
                    lstRecord.Close
                    
                    strQuery = "UPDATE DEFECT_COUNT SET "
                    strQuery = strQuery & "DEFECT_COUNT = " & intPoint_Defect_Total & " WHERE "
                    strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_DATA.DEFECT_CODE & "'"
                    
                    dbMyDB.Execute (strQuery)
                Else
                    lstRecord.Close
                    
                    strQuery = "INSERT INTO DEFECT_COUNT VALUES ("
                    strQuery = strQuery & "'" & pDEFECT_DATA.DEFECT_CODE & "', "
                    strQuery = strQuery & "1)"
                    
                    dbMyDB.Execute (strQuery)
                End If
                
                dbMyDB.Close
            End If
        Else
        End If
    End If

End Sub

Public Sub Update_Defect_Grade(pDEFECT_DATA As DEFECT_DATA_STRUCTURE)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strQuery                    As String
    Dim strCurrent_KEYID            As String
    
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    strCurrent_KEYID = RANK_OBJ.Get_Current_KEYID
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "PANELID = '" & pDEFECT_DATA.PANELID & "' AND "
        strQuery = strQuery & "KEYID = '" & strCurrent_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            strPath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        If strPath <> "" Then
            If Right(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
        End If
        If (strPath <> "") And (strFileName <> "") Then
            If Dir(strPath & strFileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strPath & strFileName)
                
                strQuery = "UPDATE DEFECT_DATA SET "
                strQuery = strQuery & "DEFECT_GRADE='" & pDEFECT_DATA.GRADE & "' WHERE "
                strQuery = strQuery & "DEFECT_NO=" & pDEFECT_DATA.DEFECT_NO & " AND "
                strQuery = strQuery & "DEFECT_CODE='" & pDEFECT_DATA.DEFECT_CODE & "'"
                
                Call dbMyDB.Execute(strQuery)
                
                dbMyDB.Close
            End If
        Else
        End If
    End If
    
End Sub

Public Function Get_Point_Defect_Total(ByVal pDEFECT_CODE As String, ByVal pPanelID As String) As Integer

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strQuery                    As String
    Dim strCurrent_KEYID            As String
    
    Dim intPoint_Defect_Total       As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    strCurrent_KEYID = RANK_OBJ.Get_Current_KEYID
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "PANELID = '" & pPanelID & "' AND "
        strQuery = strQuery & "KEYID = '" & strCurrent_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            strPath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        If Right(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strPath & strFileName)
            
            strQuery = "SELECT * FROM DEFECT_COUNT WHERE "
            strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_CODE & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                Get_Point_Defect_Total = lstRecord.Fields("DEFECT_COUNT")
            Else
                Get_Point_Defect_Total = 0
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    End If
    
End Function

Public Function Convert_Rank_Grade(pGRADE_DATA() As GRADE_DATA_STRUCTURE, pRANK_DATA As RANK_DATA_STRUCTURE, _
                                   ByVal pArray_Count As Integer, ByVal pRank As String) As String

    Dim intIndex                    As Integer
    Dim intData_Index               As Integer
    
    Dim strGrade                    As String
    
    Dim bolFind_Grade               As Boolean
    
    intData_Index = 0
    intIndex = 0
    strGrade = ""
    bolFind_Grade = False
    While bolFind_Grade = False
        intIndex = intIndex + 1
        If (pGRADE_DATA(intIndex).RANK_DIVISION = pRANK_DATA.RANK_DIVISION) And (pGRADE_DATA(intIndex).DEFECT_CODE = pRANK_DATA.DEFECT_CODE) Then
            If pGRADE_DATA(intIndex).RANK = pRank Then
                strGrade = pGRADE_DATA(intIndex).RANK
                bolFind_Grade = True
            Else
                If intIndex = pArray_Count Then
                    bolFind_Grade = True
                End If
            End If
        Else
            If intIndex = pArray_Count Then
                bolFind_Grade = True
            End If
        End If
    Wend
    
    Convert_Rank_Grade = strGrade
    
End Function

Public Function Get_Panel_Grade(pPoint_Defect_Grade As String, pDefect_Rank As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim arrGrade(1 To 9)            As String
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Grade                As String
    
    Dim intIndex                    As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "SELECT * FROM DEFECT_TYPE_PRIORITY"

        Set lstRecord = dbMyDB.OpenRecordset(strQuery)

        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                Call RANK_OBJ.Set_DEFECT_PRIOITY(lstRecord.Fields("DEFECT_TYPE"), lstRecord.Fields("DEFECT_PRIORITY"))
                lstRecord.MoveNext
            Wend
            lstRecord.Close
        End If
        dbMyDB.Close
    End If
    
    strNew_Grade = RANK_OBJ.Check_DEFECT_TYPE_PRIORITY(pDefect_Rank)
    
    If strNew_Grade = "" Then
        strNew_Grade = pPoint_Defect_Grade
    End If
    
    Get_Panel_Grade = strNew_Grade
    Call SaveLog("Get_Panel_Grade", "New Grade : " & strNew_Grade)
    
End Function

Public Function PreJudgeGradeChange1(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                     pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    
    Dim intIndex                    As Integer
    
    Dim lngX_LEFT                   As Long
    Dim lngX_Right                  As Long
    Dim lngY_Upper                  As Long
    Dim lngY_Bottom                 As Long
    
    Dim bolFind_Code                As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PreJudgeGradeChange1") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
            strQuery = "SELECT * FROM PRE_JUDGE_CHANGE_GRADE1 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
    '        strQuery = strQuery & "DATA = " & pPANEL_MES_DATA.X_TOTAL_PIXEL & " AND "
    '        strQuery = strQuery & "GATE = " & pPANEL_MES_DATA.Y_TOTAL_PIXEL & " AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                While lstRecord.EOF = False
                    lngX_LEFT = (lstRecord.Fields("DATA") * 3) / lstRecord.Fields("LIMIT_LEFT")
                    lngX_Right = (lstRecord.Fields("DATA") * 3) - ((lstRecord.Fields("DATA") * 3) / lstRecord.Fields("LIMIT_RIGHT"))
                    lngY_Upper = lstRecord.Fields("GATE") / lstRecord.Fields("LIMIT_UPPER")
                    lngY_Bottom = lstRecord.Fields("GATE") - (lstRecord.Fields("GATE") / lstRecord.Fields("LIMIT_BOTTOM"))
                    If (lngX_LEFT <= CLng(pDEFECT_DATA.DATA_ADDRESS(1))) And (CLng(pDEFECT_DATA.DATA_ADDRESS(1)) <= lngX_Right) Then
                        If (lngY_Upper <= CLng(pDEFECT_DATA.GATE_ADDRESS(1))) And (CLng(pDEFECT_DATA.GATE_ADDRESS(1)) <= lngY_Bottom) Then
                            bolFind_Code = False
                            For intIndex = 1 To 10
                                If lstRecord.Fields("DEFECT_CODE" & intIndex) = pDEFECT_DATA.DEFECT_CODE Then
                                    bolFind_Code = True
                                End If
                            Next intIndex
                            If bolFind_Code = True Then
                                strNew_Judge = lstRecord.Fields("NEW_GRADE")
                            End If
                        End If
                    End If
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PreJudgeGradeChange1 = strNew_Judge
    Call SaveLog("PreJudgeGradeChange1", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PreJudgeGradeChange1", ErrMsg)
    
    PreJudgeGradeChange1 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PreJudgeGradeChange2(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                     pDEFECT_DATA() As DEFECT_DATA_STRUCTURE, ByVal pDEFECT_COUNT As Integer) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String

    Dim lngX_LEFT                   As Long
    Dim lngX_GAP                    As Long
    
    Dim intIndex                    As Integer
    Dim intLeft_Defect_Count        As Integer
    Dim intRight_Defect_Count       As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PreJudgeGradeChange2") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
            strQuery = "SELECT * FROM PRE_JUDGE_CHANGE_GRADE2 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                intLeft_Defect_Count = 0
                intRight_Defect_Count = 0
                lstRecord.MoveFirst
                While lstRecord.EOF = False
                    lngX_GAP = (lstRecord.Fields("DATA") * 3) / lstRecord.Fields("TOTAL_DIVISION")
                    lngX_LEFT = lngX_GAP * lstRecord.Fields("DEVIDE_DIVISION")
                    For intIndex = 1 To pDEFECT_COUNT
                        If (pDEFECT_DATA(intIndex).DEFECT_CODE = lstRecord.Fields("DEFECT_CODE1")) Or _
                           (pDEFECT_DATA(intIndex).DEFECT_CODE = lstRecord.Fields("DEFECT_CODE2")) Or _
                           (pDEFECT_DATA(intIndex).DEFECT_CODE = lstRecord.Fields("DEFECT_CODE3")) Then
                            If (Trim(lstRecord.Fields("PRE_GRADE")) = "") Or (lstRecord.Fields("PRE_GRADE") = pPre_Judge) Then
                                If (0 <= CLng(pDEFECT_DATA(intIndex).DATA_ADDRESS(1))) And (CLng(pDEFECT_DATA(intIndex).DATA_ADDRESS(1)) <= lngX_LEFT) Then
                                    intLeft_Defect_Count = intLeft_Defect_Count + CInt(frmJudge.flxDefect_List.TextMatrix(intIndex, 10))
                                Else
                                    intRight_Defect_Count = intRight_Defect_Count + CInt(frmJudge.flxDefect_List.TextMatrix(intIndex, 10))
                                End If
                            End If
                        End If
                    Next intIndex
                    If (lstRecord.Fields("LIMIT_COUNT") < intLeft_Defect_Count) Or (lstRecord.Fields("LIMIT_COUNT") < intRight_Defect_Count) Then
                        strNew_Judge = lstRecord.Fields("NEW_GRADE")
                    End If
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PreJudgeGradeChange2 = strNew_Judge
    Call SaveLog("PreJudgeGradeChange2", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PreJudgeGradeChange2", ErrMsg)
    
    PreJudgeGradeChange2 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PreJudgeGradeChange3(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                     pDEFECT_DATA As DEFECT_DATA_STRUCTURE, ByVal pPoint_Defect_Rank As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    Dim strDEFECT_TYPE              As String
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PreJudgeGradeChange3") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM PRE_JUDGE_CHANGE_GRADE3 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "' AND "
            strQuery = strQuery & "POINT_DEFECT_GRADE = '" & RANK_OBJ.Get_GRADE_by_DEFECT_TYPE("D") & "' AND "
            strQuery = strQuery & "DEFECT_TYPE = '" & strDEFECT_TYPE & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strNew_Judge = lstRecord.Fields("NEW_GRADE")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PreJudgeGradeChange3 = strNew_Judge
    Call SaveLog("PreJudgeGradeChange3", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PreJudgeGradeChange3", ErrMsg)
    
    PreJudgeGradeChange3 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PostJudgeOtherRule1(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strDEFECT_TYPE              As String
    Dim strNew_Judge                As String

    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PostJudgeOtherRule1") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM POST_JUDGE_OTHER_RULE1 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "'"
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                If (lstRecord.Fields("SOURCE_DEFECT_CODE") = " ") Or (lstRecord.Fields("SOURCE_DEFECT_CODE") = pDEFECT_DATA.DEFECT_CODE) Then
                    If (lstRecord.Fields("CELL_REPAIR_GRADE") = " ") Or (lstRecord.Fields("CELL_REPAIR_GRADE") = pPANEL_MES_DATA.CELL_REPAIR_JUDGE_GRADE) Then
                        If (lstRecord.Fields("PRE_GRADE") = " ") Or (lstRecord.Fields("PRE_GRADE") = pPre_Judge) Then
                            strNew_Judge = lstRecord.Fields("NEW_GRADE")
                        End If
                    End If
                End If
            Else
                strNew_Judge = pPre_Judge
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PostJudgeOtherRule1 = strNew_Judge
    Call SaveLog("PostJudgeOtherRule1", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PostJudgeOtherRule1", ErrMsg)
    
    PostJudgeOtherRule1 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PostJudgeOtherRule2(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strDEFECT_TYPE              As String
    Dim strNew_Judge                As String

    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PostJudgeOtherRule2") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM POST_JUDGE_OTHER_RULE2 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "' AND "
            strQuery = strQuery & "DEFECT_CODE = '" & pDEFECT_DATA.DEFECT_CODE & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strNew_Judge = lstRecord.Fields("NEW_GRADE")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PostJudgeOtherRule2 = strNew_Judge
    Call SaveLog("PostJudgeOtherRule2", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PostJudgeOtherRule2", ErrMsg)
    
    PostJudgeOtherRule2 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PostJudgeOtherRule3(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strDEFECT_TYPE              As String
    Dim strNew_Judge                As String

    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PostJudgeOtherRule3") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM POST_JUDGE_OTHER_RULE3 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_LOSS_CODE = '" & pPANEL_MES_DATA.LIGHT_ON_REASON_CODE & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strNew_Judge = lstRecord.Fields("NEW_GRADE")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PostJudgeOtherRule3 = strNew_Judge
    Call SaveLog("PostJudgeOtherRule3", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PostJudgeOtherRule3", ErrMsg)
    
    PostJudgeOtherRule3 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PostJudgeGradeChange1(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim typDEFECT_DATA              As DEFECT_DATA_STRUCTURE
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strDEFECT_TYPE              As String
    Dim strNew_Judge                As String

    Dim intDefect_Index             As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PostJudgeGradeChange1") = "E" Then
        intDefect_Index = RANK_OBJ.Get_DEFECT_INDEX_by_DEFECT_TYPE("L")
        
        If intDefect_Index > 0 Then
            With typDEFECT_DATA
                If RANK_OBJ.Get_DEFECT_DATA_by_Index(intDefect_Index, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                     .GRADE, .RANK, .COLOR, .GRAY_LEVEL, .ACCUMULATION) = False Then
                    Call SaveLog("PostJudgeGradeChange1", intDefect_Index & "'s defect data is not found.")
                End If
            End With
            strDB_Path = App.PATH & "\DB\"
            strDB_FileName = "STANDARD_INFO.mdb"
            strNew_Judge = pPre_Judge
            strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
                strQuery = "SELECT * FROM POST_JUDGE_GRADE_CHANGE1 WHERE "
                strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
                strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
                strQuery = strQuery & "ABNORMAL_TFT = '" & pPANEL_MES_DATA.ABNORMAL_TFT & "'"
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    lstRecord.MoveFirst
                    If (typDEFECT_DATA.DATA_ADDRESS(1) <= lstRecord.Fields("DATA_LINE")) Or (typDEFECT_DATA.DATA_ADDRESS(2) <= lstRecord.Fields("DATA_LINE")) Then
                        strNew_Judge = lstRecord.Fields("LEFT_GRADE")
                    Else
                        strNew_Judge = lstRecord.Fields("RIGHT_GRADE")
                    End If
                Else
                    strNew_Judge = pPre_Judge
                End If
                lstRecord.Close
                
                dbMyDB.Close
            End If
        Else
            strNew_Judge = pPre_Judge
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PostJudgeGradeChange1 = strNew_Judge
    Call SaveLog("PostJudgeGradeChange1", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PostJudgeGradeChange1", ErrMsg)
    
    PostJudgeGradeChange1 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function PostJudgeGradeChange2(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strDEFECT_TYPE              As String
    Dim strNew_Judge                As String

    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("PostJudgeGradeChange2") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM POST_JUDGE_GRADE_CHANGE2 WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "' AND "
            strQuery = strQuery & "TFT_REPAIR_GRADE = '" & pPANEL_MES_DATA.TFT_REPAIR_GRADE & "' AND "
            strQuery = strQuery & "CELL_LINE_RESCUE_FLAG = '" & pPANEL_MES_DATA.CELL_LINE_RESCUE_FLAG & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strNew_Judge = lstRecord.Fields("NEW_GRADE")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    PostJudgeGradeChange2 = strNew_Judge
    Call SaveLog("PostJudgeGradeChange2", "New Grade : " & strNew_Judge)
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("PostJudgeGradeChange2", ErrMsg)

    PostJudgeGradeChange2 = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function CheckPanelIDChangeGrade(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    Dim strDEFECT_TYPE              As String
    
    Dim intIndex                    As Integer
    Dim intFileIndex                As Integer
    
    Dim bolNo_Change_Grade          As Boolean
    Dim bolEnable                   As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    'Control Item data processing logic
    CheckPanelIDChangeGrade = pPre_Judge
    For intFileIndex = 1 To 5
        If EQP.Get_Control_Data("CheckPanelIDChangeGrade" & intFileIndex) = "E" Then
            strDB_Path = App.PATH & "\DB\"
            strDB_FileName = "STANDARD_INFO.mdb"
            strNew_Judge = pPre_Judge
            strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
                strQuery = "SELECT * FROM CHECK_PANELID_CHANGE_GRADE WHERE "
                strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
                strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
                strQuery = strQuery & "PANELID = '" & pPANEL_MES_DATA.PANELID & "' AND "
                strQuery = strQuery & "FILE_NAME = '" & "CheckPanelIDChangeGrade" & intFileIndex & "'"
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    lstRecord.MoveFirst
                    bolNo_Change_Grade = False
                    For intIndex = 1 To 10
                        If lstRecord.Fields("NO_CHANGE_GRADE" & intIndex) = pPre_Judge Then
                            bolNo_Change_Grade = True
                        End If
                    Next intIndex
                    If bolNo_Change_Grade = False Then
                        strNew_Judge = lstRecord.Fields("NEW_GRADE")
                        CheckPanelIDChangeGrade = strNew_Judge
                        Call SaveLog("CheckPanelIDChangeGrade" & intFileIndex, "New Grade : " & strNew_Judge)
                    End If
                Else
                    Call SaveLog("CheckPanelIDChangeGrade" & intFileIndex, "Record does not exist.")
                End If
                lstRecord.Close
                
                dbMyDB.Close
            End If
        Else
            Call SaveLog("CheckPanelIDChangeGrade" & intFileIndex, "Control value : D")
        End If
    Next intFileIndex
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("CheckPanelIDChangeGrade", ErrMsg)
    
    CheckPanelIDChangeGrade = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function ChangeGrade(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    Dim strDEFECT_TYPE              As String
    Dim strMachineID                As String
    Dim strUnitID                   As String
    
    Dim intIndex                    As Integer
    
    Dim bolFind_UnitID              As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("ChangeGrade") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        strDEFECT_TYPE = Mid(pDEFECT_DATA.DEFECT_CODE, 2, 1)
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
            strQuery = "SELECT * FROM CHANGE_GRADE WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "' AND "
            strMachineID = ENV.Get_Current_Machine_Name
            strUnitID = Right(strMachineID, 3)
            strMachineID = Left(strMachineID, 5)
            strQuery = strQuery & "MACHINE_TYPE = '" & strMachineID & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                bolFind_UnitID = False
                For intIndex = 1 To 6
                    If lstRecord.Fields("MACHINE_ID" & intIndex) = strUnitID Then
                        bolFind_UnitID = True
                    End If
                Next intIndex
                If bolFind_UnitID = True Then
                    strNew_Judge = lstRecord.Fields("NEW_GRADE")
                End If
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    ChangeGrade = strNew_Judge
    Call SaveLog("ChangeGrade", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("ChangeGrade", ErrMsg)
    
    ChangeGrade = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function ChangeGradeByDefectCode(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("ChangeGradeByDefectCode") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM CHANGE_GRADE_DEFECT_CODE WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "' AND "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "' AND "
            strQuery = strQuery & "LIGHT_ON_REASON_CODE = '" & pPANEL_MES_DATA.LIGHT_ON_REASON_CODE & "' AND "
            strQuery = strQuery & "DRIVE_TYPE = '" & frmMain.flxEQ_Information.TextMatrix(1, 1) & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                strNew_Judge = lstRecord.Fields("NEW_GRADE")
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    ChangeGradeByDefectCode = strNew_Judge
    Call SaveLog("ChangeGradeByDefectCode", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("ChangeGradeByDefectCode", ErrMsg)
    
    ChangeGradeByDefectCode = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function RepairPointTimes(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    
    Dim intIndex                    As Integer
    
    Dim bolFind_Defect_Code         As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("RepairPointTimes") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM REPAIR_POINT_TIMES WHERE "
            strQuery = strQuery & "PFCD = '" & Mid(pCST_MES_DATA.PFCD, 3, 5) & "' AND "
            strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                While lstRecord.EOF = False
                    bolFind_Defect_Code = False
                    If pPANEL_MES_DATA.REPAIR_REWORK_COUNT = "" Then
                        pPANEL_MES_DATA.REPAIR_REWORK_COUNT = "0"
                    End If
                    
                    Select Case pPANEL_MES_DATA.REPAIR_REWORK_COUNT
                    
                    Case "0":
                    
                      If CInt(pPANEL_MES_DATA.REPAIR_REWORK_COUNT) < lstRecord.Fields("REPAIR_REWORK_COUNT") Then
                        For intIndex = 1 To 5
                            If lstRecord.Fields("DEFECT_CODE" & intIndex) = pDEFECT_DATA.DEFECT_CODE Then
                                bolFind_Defect_Code = True
                            End If
                        Next intIndex
                        If bolFind_Defect_Code = True Then
                            strNew_Judge = lstRecord.Fields("NEW_GRADE")
                        End If
                     End If
                     
                     If pCST_MES_DATA.PROCESS_NUM = "3650" Then
                       If (pDEFECT_DATA.DEFECT_CODE = "CDRBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDGBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDBBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDBTD") Or (pDEFECT_DATA.DEFECT_CODE = "CDBDD") Or (pDEFECT_DATA.DEFECT_CODE = "CDLEK") Or (pDEFECT_DATA.DEFECT_CODE = "CDBTT") Then
                        strNew_Judge = "RC"
                       End If
                    End If
                     
'Lucas 2012.02.22===================================For Repair rework_count
                    Case "2":
                          If (pCST_MES_DATA.PROCESS_NUM = "3650") Then
'Lucas Ver.0.9.29 2012.05.30=======================For CUT Test
                             If (pDEFECT_DATA.DEFECT_CODE = "CLDDK") Or (pDEFECT_DATA.DEFECT_CODE = "CLDWK") Or (pDEFECT_DATA.DEFECT_CODE = "CLDBT") Then
                              strNew_Judge = "RP"
'Lucas Ver.0.9.29 2012.05.30=======================For CUT Test
                             Else
                              strNew_Judge = "NG"
                             End If
                          Else
                            If pCST_MES_DATA.PROCESS_NUM = "4650" Then
                            strNew_Judge = "S "
                            End If
                          End If
 '===============================End
 'Lucas 2012.04.13===================================For Repair case "RP"
                    Case "3":
                          If (pCST_MES_DATA.PROCESS_NUM = "3650") Then
                           strNew_Judge = "RP"
                           Else
                            If pCST_MES_DATA.PROCESS_NUM = "4650" Then
                            strNew_Judge = "RT"
                          End If
                        End If
                        
 'Lucas 2012.05.06===================================For CRP "Y" case "RC"
                    Case "1":
                         If pCST_MES_DATA.PROCESS_NUM = "3650" Then
                            If (pDEFECT_DATA.DEFECT_CODE = "CDRBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDGBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDBBP") Or (pDEFECT_DATA.DEFECT_CODE = "CDBTD") Or (pDEFECT_DATA.DEFECT_CODE = "CDBDD") Or (pDEFECT_DATA.DEFECT_CODE = "CDLEK") Or (pDEFECT_DATA.DEFECT_CODE = "CDBTT") Then
                             strNew_Judge = "RC"
                            End If
                         End If
                   End Select
 '===============================End
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    RepairPointTimes = strNew_Judge
    Call SaveLog("RepairPointTimes", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("RepairPointTimes", ErrMsg)
    
    RepairPointTimes = strNew_Judge
    
    dbMyDB.Close
    
End Function


Public Function FlagChangeGrade(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE, pJOB_DATA As JOB_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim arrResult(1 To 52)          As Boolean
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String
    
    Dim intIndex                    As Integer
    
    Dim bolFind_False               As Boolean
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("FlagChangeGrade") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM FLAG_CHANGE_GRADE WHERE "
            strQuery = strQuery & "PFCD = '" & pCST_MES_DATA.PFCD & "' And "
            strQuery = strQuery & "PRE_GRADE = '" & pPre_Judge & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                For intIndex = 1 To 52
                    arrResult(intIndex) = True
                Next intIndex
                
                While lstRecord.EOF = False
                    With lstRecord
                        If (.Fields("DEST_FAB") = " ") Or (.Fields("DEST_FAB") = pCST_MES_DATA.DESTINATION_FAB) Or (.Fields("DEST_FAB") = "") Then
                            arrResult(1) = True
                        Else
                            arrResult(1) = False
                        End If
                        
                        If (.Fields("RMANO") = " ") Or (.Fields("RMANO") = pCST_MES_DATA.RMANO) Or (.Fields("RMANO") = "") Then
                            arrResult(2) = True
                        Else
                            arrResult(2) = False
                        End If
                        
                        If (.Fields("OQCNO") = " ") Or (.Fields("OQCNO") = pCST_MES_DATA.OQCNO) Or (.Fields("OQCNO") = "") Then
                            arrResult(3) = True
                        Else
                            arrResult(3) = False
                        End If
                        
                        If (.Fields("PANELID") = " ") Or (.Fields("PANELID") = pPANEL_MES_DATA.PANELID) Or (.Fields("PANELID") = "") Then
                            arrResult(4) = True
                        Else
                            arrResult(4) = False
                        End If
                        
                        If (.Fields("LIGHT_ON_PANEL_GRADE") = " ") Or (.Fields("LIGHT_ON_PANEL_GRADE") = pPANEL_MES_DATA.LIGHT_ON_PANEL_GRADE) Or (.Fields("LIGHT_ON_PANEL_GRADE") = "") Then
                            arrResult(5) = True
                        Else
                            arrResult(5) = False
                        End If
                        
                        If (.Fields("LIGHT_ON_REASON_CODE") = " ") Or (.Fields("LIGHT_ON_REASON_CODE") = pPANEL_MES_DATA.LIGHT_ON_REASON_CODE) Or (.Fields("LIGHT_ON_REASON_CODE") = "") Then
                            arrResult(6) = True
                        Else
                            arrResult(6) = False
                        End If
                        
                        If (.Fields("CELL_LINE_RESCUE_FLAG") = " ") Or (.Fields("CELL_LINE_RESCUE_FLAG") = pPANEL_MES_DATA.CELL_LINE_RESCUE_FLAG) Or (.Fields("CELL_LINE_RESCUE_FLAG") = "") Then
                            arrResult(7) = True
                        Else
                            arrResult(7) = False
                        End If
                        
                        If (.Fields("CELL_REPAIR_JUDGE_GRADE") = " ") Or (.Fields("CELL_REPAIR_JUDGE_GRADE") = pPANEL_MES_DATA.CELL_REPAIR_JUDGE_GRADE) Or (.Fields("CELL_REPAIR_JUDGE_GRADE") = "") Then
                            arrResult(8) = True
                        Else
                            arrResult(8) = False
                        End If
                        
                         If (.Fields("TFT_REPAIR_GRADE") = " ") Or (.Fields("TFT_REPAIR_GRADE") = pPANEL_MES_DATA.TFT_REPAIR_GRADE) Or (.Fields("TFT_REPAIR_GRADE") = "") Then
                            arrResult(9) = True
                        Else
                            arrResult(9) = False
                        End If
                        
                        If (.Fields("CF_PANELID") = " ") Or (.Fields("CF_PANELID") = pPANEL_MES_DATA.CF_PANELID) Or (.Fields("CF_PANELID") = "") Then
                            arrResult(10) = True
                        Else
                            arrResult(10) = False
                        End If
                        
                        If (.Fields("CF_PANEL_OX_INFORMATION") = " ") Or (.Fields("CF_PANEL_OX_INFORMATION") = pPANEL_MES_DATA.CF_PANEL_OX_INFORMATION) Or (.Fields("CF_PANEL_OX_INFORMATION") = "") Then
                            arrResult(11) = True
                        Else
                            arrResult(11) = False
                        End If
                        
                        If (.Fields("PANEL_OWNER_TYPE") = " ") Or (.Fields("PANEL_OWNER_TYPE") = pPANEL_MES_DATA.PANEL_OWNER_TYPE) Or (.Fields("PANEL_OWNER_TYPE") = "") Then
                            arrResult(12) = True
                        Else
                            arrResult(12) = False
                        End If
                        
                        If (.Fields("ABNORMAL_CF") = " ") Or (.Fields("ABNORMAL_CF") = pPANEL_MES_DATA.ABNORMAL_CF) Or (.Fields("ABNORMAL_CF") = "") Then
                            arrResult(13) = True
                        Else
                            arrResult(13) = False
                        End If
                        
                        If (.Fields("ABNORMAL_TFT") = " ") Or (.Fields("ABNORMAL_TFT") = pPANEL_MES_DATA.ABNORMAL_TFT) Or (.Fields("ABNORMAL_TFT") = "") Then
                            arrResult(14) = True
                        Else
                            arrResult(14) = False
                        End If
                        
                        If (.Fields("ABNORMAL_LCD") = " ") Or (.Fields("ABNORMAL_LCD") = pPANEL_MES_DATA.ABNORMAL_LCD) Or (.Fields("ABNORMAL_LCD") = "") Then
                            arrResult(15) = True
                        Else
                            arrResult(15) = False
                        End If
                        
                        If (.Fields("GROUPID") = " ") Or (.Fields("GROUPID") = pPANEL_MES_DATA.GROUP_ID) Or (.Fields("GROUPID") = "") Then
                            arrResult(16) = True
                        Else
                            arrResult(16) = False
                        End If
                        
                        If (.Fields("REPAIR_REWORK_COUNT") = " ") Or (.Fields("REPAIR_REWORK_COUNT") = pPANEL_MES_DATA.REPAIR_REWORK_COUNT) Or (.Fields("REPAIR_REWORK_COUNT") = "") Then
                            arrResult(17) = True
                        Else
                            arrResult(17) = False
                        End If
                        
                        If (.Fields("POLARIZER_REWORK_COUNT") = " ") Or (.Fields("POLARIZER_REWORK_COUNT") = pPANEL_MES_DATA.POLARIZER_REWORK_COUNT) Or (.Fields("POLARIZER_REWORK_COUNT") = "") Then
                            arrResult(18) = True
                        Else
                            arrResult(18) = False
                        End If
                        
                        If (.Fields("X_TOTAL_PIXEL") = " ") Or (.Fields("X_TOTAL_PIXEL") = pPANEL_MES_DATA.X_TOTAL_PIXEL) Or (.Fields("X_TOTAL_PIXEL") = "") Then
                            arrResult(19) = True
                        Else
                            arrResult(19) = False
                        End If
                        
                        If (.Fields("Y_TOTAL_PIXEL") = " ") Or (.Fields("Y_TOTAL_PIXEL") = pPANEL_MES_DATA.Y_TOTAL_PIXEL) Or (.Fields("Y_TOTAL_PIXEL") = "") Then
                            arrResult(20) = True
                        Else
                            arrResult(20) = False
                        End If
                        
                        If (.Fields("LCD_Q_TAB_LOT_GROUPID") = " ") Or (.Fields("LCD_Q_TAB_LOT_GROUPID") = pPANEL_MES_DATA.LCD_Q_TAP_LOT_GROUPID) Or (.Fields("LCD_Q_TAB_LOT_GROUPID") = "") Then
                            arrResult(21) = True
                        Else
                            arrResult(21) = False
                        End If
                        
                        If (.Fields("SK_FLAG") = " ") Or (.Fields("SK_FLAG") = pPANEL_MES_DATA.SK_FLAG) Or (.Fields("SK_FLAG") = "") Then
                            arrResult(22) = True
                        Else
                            arrResult(22) = False
                        End If
                        
                        If (.Fields("CF_R_DEFECT_CODE") = " ") Or (.Fields("CF_R_DEFECT_CODE") = pPANEL_MES_DATA.CF_R_DEFECT_CODE) Or (.Fields("CF_R_DEFECT_CODE") = "") Then
                            arrResult(23) = True
                        Else
                            arrResult(23) = False
                        End If
                        
                        If (.Fields("ODK_AK_FLAG") = " ") Or (.Fields("ODK_AK_FLAG") = pPANEL_MES_DATA.ODK_AK_FLAG) Or (.Fields("ODK_AK_FLAG") = "") Then
                            arrResult(24) = True
                        Else
                            arrResult(24) = False
                        End If
                        
                        If (.Fields("BPAM_REWORK_FLAG") = " ") Or (.Fields("BPAM_REWORK_FLAG") = pPANEL_MES_DATA.BPAM_REWORK_FLAG) Or (.Fields("BPAM_REWORK_FLAG") = "") Then
                            arrResult(25) = True
                        Else
                            arrResult(25) = False
                        End If
                        
                        If (.Fields("LCD_BRIGHT_DOT_FLAG") = " ") Or (.Fields("LCD_BRIGHT_DOT_FLAG") = pPANEL_MES_DATA.LCD_BRIGHT_DOT_FLAG) Or (.Fields("LCD_BRIGHT_DOT_FLAG") = "") Then
                            arrResult(26) = True
                        Else
                            arrResult(26) = False
                        End If
                        
                        If (.Fields("CF_PS_HIGHT_ERR_FLAG") = " ") Or (.Fields("CF_PS_HIGHT_ERR_FLAG") = pPANEL_MES_DATA.CF_PS_HEIGHT_ERR_FLAG) Or (.Fields("CF_PS_HIGHT_ERR_FLAG") = "") Then
                            arrResult(27) = True
                        Else
                            arrResult(27) = False
                        End If
                        
                        If (.Fields("PI_INSPECTION_NG_FLAG") = " ") Or (.Fields("PI_INSPECTION_NG_FLAG") = pPANEL_MES_DATA.PI_INSPECTION_NG_FLAG) Or (.Fields("PI_INSPECTION_NG_FLAG") = "") Then
                            arrResult(28) = True
                        Else
                            arrResult(28) = False
                        End If
                        
                        If (.Fields("PI_OVER_BAKE_FLAG") = " ") Or (.Fields("PI_OVER_BAKE_FLAG") = pPANEL_MES_DATA.PI_OVER_BAKE_FLAG) Or (.Fields("PI_OVER_BAKE_FLAG") = "") Then
                            arrResult(29) = True
                        Else
                            arrResult(29) = False
                        End If
                        
                        If (.Fields("PI_OVER_Q_TIME_FLAG") = " ") Or (.Fields("PI_OVER_Q_TIME_FLAG") = pPANEL_MES_DATA.PI_OVER_Q_TIME_FLAG) Or (.Fields("PI_OVER_Q_TIME_FLAG") = "") Then
                            arrResult(30) = True
                        Else
                            arrResult(30) = False
                        End If
                        
                        If (.Fields("ODF_OVER_BAKE_FLAG") = " ") Or (.Fields("ODF_OVER_BAKE_FLAG") = pPANEL_MES_DATA.ODF_OVER_BAKE_FLAG) Or (.Fields("ODF_OVER_BAKE_FLAG") = "") Then
                            arrResult(31) = True
                        Else
                            arrResult(31) = False
                        End If
                        
                        If (.Fields("ODF_OVER_Q_TIME_FLAG") = " ") Or (.Fields("ODF_OVER_Q_TIME_FLAG") = pPANEL_MES_DATA.ODF_OVER_Q_TIME_FLAG) Or (.Fields("ODF_OVER_Q_TIME_FLAG") = "") Then
                            arrResult(32) = True
                        Else
                            arrResult(32) = False
                        End If
                        
                        If (.Fields("HVA_OVER_BAKE_FLAG") = " ") Or (.Fields("HVA_OVER_BAKE_FLAG") = pPANEL_MES_DATA.HVA_OVER_BAKE_FLAG) Or (.Fields("HVA_OVER_BAKE_FLAG") = "") Then
                            arrResult(33) = True
                        Else
                            arrResult(33) = False
                        End If
                        
                        If (.Fields("HVA_OVER_Q_TIME_FLAG") = " ") Or (.Fields("HVA_OVER_Q_TIME_FLAG") = pPANEL_MES_DATA.HVA_OVER_Q_TIME_FLAG) Or (.Fields("HVA_OVER_Q_TIME_FLAG") = "") Then
                            arrResult(34) = True
                        Else
                            arrResult(34) = False
                        End If
                        
                        If (.Fields("SEAL_INSPECTION_FLAG") = " ") Or (.Fields("SEAL_INSPECTION_FLAG") = pPANEL_MES_DATA.SEAL_INSPECTION_FLAG) Or (.Fields("SEAL_INSPECTION_FLAG") = "") Then
                            arrResult(35) = True
                        Else
                            arrResult(35) = False
                        End If
                        
                        If (.Fields("ODF_CHECKER_FLAG") = " ") Or (.Fields("ODF_CHECKER_FLAG") = pPANEL_MES_DATA.ODF_CHECKER_FLAG) Or (.Fields("ODF_CHECKER_FLAG") = "") Then
                            arrResult(36) = True
                        Else
                            arrResult(36) = False
                        End If
                        
                        If (.Fields("ODF_DOOR_OPEN_FLAG") = " ") Or (.Fields("ODF_DOOR_OPEN_FLAG") = pPANEL_MES_DATA.ODF_DOOR_OPEN_FLAG) Or (.Fields("ODF_DOOR_OPEN_FLAG") = "") Then
                            arrResult(37) = True
                        Else
                            arrResult(37) = False
                        End If
                        
                        If (.Fields("JOB_JUDGE") = " ") Or (.Fields("JOB_JUDGE") = pJOB_DATA.JOB_JUDGE) Or (.Fields("JOB_JUDGE") = "") Then
                            arrResult(38) = True
                        Else
                            arrResult(38) = False
                        End If
                        
                        If (.Fields("JOB_GRADE") = " ") Or (.Fields("JOB_GRADE") = pJOB_DATA.JOB_GRADE) Or (.Fields("JOB_GRADE") = "") Then
                            arrResult(39) = True
                        Else
                            arrResult(39) = False
                        End If
                        
                        If (.Fields("BURR_CHECK_JUDGE") = " ") Or (.Fields("BURR_CHECK_JUDGE") = pJOB_DATA.BURR_CHECK_JUDGE) Or (.Fields("BURR_CHECK_JUDGE") = "") Then
                            arrResult(40) = True
                        Else
                            arrResult(40) = False
                        End If
                        
                        If (.Fields("BEVELING_JUDGE") = " ") Or (.Fields("BEVELING_JUDGE") = pJOB_DATA.BEVELING_JUDGE) Or (.Fields("BEVELING_JUDGE") = "") Then
                            arrResult(41) = True
                        Else
                            arrResult(41) = False
                        End If
                        
                        If (.Fields("CLEANER_M_PORT_JUDGE") = " ") Or (.Fields("CLEANER_M_PORT_JUDGE") = pJOB_DATA.CLEANER_M_PORT_JUDGE) Or (.Fields("CLEANER_M_PORT_JUDGE") = "") Then
                            arrResult(42) = True
                        Else
                            arrResult(42) = False
                        End If
                        
                        If (.Fields("TEST_CV_JUDGE") = " ") Or (.Fields("TEST_CV_JUDGE") = pJOB_DATA.TEST_CV_JUDGE) Or (.Fields("TEST_CV_JUDGE") = "") Then
                            arrResult(43) = True
                        Else
                            arrResult(43) = False
                        End If
                        
                        If (.Fields("SAMPLING_SLOT_FLAG") = " ") Or (.Fields("SAMPLING_SLOT_FLAG") = pJOB_DATA.SAMPLING_SLOT_FLAG) Or (.Fields("SAMPLING_SLOT_FLAG") = "") Then
                            arrResult(44) = True
                        Else
                            arrResult(44) = False
                        End If
                        
                        If (.Fields("PROCESS_INPUT_FLAG") = " ") Or (.Fields("PROCESS_INPUT_FLAG") = pJOB_DATA.PROCESS_INPUT_FLAG) Or (.Fields("PROCESS_INPUT_FLAG") = "") Then
                            arrResult(45) = True
                        Else
                            arrResult(45) = False
                        End If
                        
                        If (.Fields("NEED_GRINDING_FLAG") = " ") Or (.Fields("NEED_GRINDING_FLAG") = pJOB_DATA.NEED_GRINDING_FLAG) Or (.Fields("NEED_GRINDING_FLAG") = "") Then
                            arrResult(46) = True
                        Else
                            arrResult(46) = False
                        End If
                        
                        If (.Fields("MISALIGNMENT_FLAG") = " ") Or (.Fields("MISALIGNMENT_FLAG") = pJOB_DATA.MISALIGNMENT_FLAG) Or (.Fields("MISALIGNMENT_FLAG") = "") Then
                            arrResult(47) = True
                        Else
                            arrResult(47) = False
                        End If
                        
                        If (.Fields("SMALL_MULTI_PANEL_FLAG") = " ") Or (.Fields("SMALL_MULTI_PANEL_FLAG") = pJOB_DATA.SMALL_MULTI_PANEL_FLAG) Or (.Fields("SMALL_MULTI_PANEL_FLAG") = "") Then
                            arrResult(48) = True
                        Else
                            arrResult(48) = False
                        End If
                        
                        If (.Fields("NO_MATCH_GLASS_IN_BC_FLAG") = " ") Or (.Fields("NO_MATCH_GLASS_IN_BC_FLAG") = pJOB_DATA.NO_MATCH_GLASS_IN_BC_FLAG) Or (.Fields("NO_MATCH_GLASS_IN_BC_FLAG") = "") Then
                            arrResult(49) = True
                        Else
                            arrResult(49) = False
                        End If
                        
                        If (.Fields("ABNORMAL_FLAG_CODE") = " ") Or (.Fields("ABNORMAL_FLAG_CODE") = pJOB_DATA.ABNORMAL_FLAG_CODE) Or (.Fields("ABNORMAL_FLAG_CODE") = "") Then
                            arrResult(50) = True
                        Else
                            arrResult(50) = False
                        End If
                        
                        If (.Fields("PANEL_NG_FLAG") = " ") Or (.Fields("PANEL_NG_FLAG") = pJOB_DATA.PANEL_NG_FLAG) Or (.Fields("PANEL_NG_FLAG") = "") Then
                            arrResult(51) = True
                        Else
                            arrResult(51) = False
                        End If
                        
                        If (.Fields("CUT_FLAG") = " ") Or (.Fields("CUT_FLAG") = pJOB_DATA.CUT_FLAG) Or (.Fields("CUT_FLAG") = "") Then
                            arrResult(52) = True
                        Else
                            arrResult(52) = False
                        End If
                        
                        bolFind_False = False
                        For intIndex = 1 To 52
                            If arrResult(intIndex) = False Then
                                bolFind_False = True
                            End If
                        Next intIndex
                        
                        If bolFind_False = False Then
                            strNew_Judge = .Fields("NEW_GRADE")
                        End If
                    End With
                    
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            
            
            
'            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
'
'            If lstRecord.EOF = False Then
'                lstRecord.MoveFirst
'                strNew_Judge = lstRecord.Fields("NEW_GRADE")
'            End If
'            lstRecord.Close
            
            dbMyDB.Close
        End If
    Else
        strNew_Judge = pPre_Judge
    End If
    
    FlagChangeGrade = strNew_Judge
    Call SaveLog("FlagChangeGrade", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("FlagChangeGrade", ErrMsg)
    
    FlagChangeGrade = strNew_Judge
    
    dbMyDB.Close
    
End Function

Public Function SKChange(ByVal pPre_Judge As String, pCST_MES_DATA As CST_INFO_ELEMENTS, pPANEL_MES_DATA As PANEL_INFO_ELEMENTS, _
                                    pDEFECT_DATA As DEFECT_DATA_STRUCTURE) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strNew_Judge                As String

    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    If EQP.Get_Control_Data("FlagChangeGrade") = "E" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "STANDARD_INFO.mdb"
        strNew_Judge = pPre_Judge
        If pubJOB_INFO.SK_FLAG = "0" Then
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
                strQuery = "SELECT * FROM SK_CHANGE WHERE "
                strQuery = strQuery & "MACHINE_NAME = '" & frmMain.flxEQ_Information.TextMatrix(5, 1) & "' AND "
                strQuery = strQuery & "PROCESSNUM = '" & pCST_MES_DATA.PROCESS_NUM & "'"
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
                If lstRecord.EOF = False Then
                    If RANK_OBJ.Increase_SK_Panel_Count = True Then
                        strNew_Judge = lstRecord.Fields("NEW_GRADE")
                        pubJOB_INFO.SK_FLAG = lstRecord.Fields("NEW_SK_FLAG")
                        Call Save_MES_Data(pubCST_INFO, pubPANEL_INFO, pubJOB_INFO, pubSHARE_INFO)
                        Call Get_MES_Data(pubCST_INFO, pubPANEL_INFO, pubJOB_INFO, pubSHARE_INFO)
                        lstRecord.Close
                        
                        Call Update_SK_Flag
                    End If
                Else
                    lstRecord.Close
                End If
                
                dbMyDB.Close
            End If
        End If
    Else
        strNew_Judge = pPre_Judge
    End If

    SKChange = strNew_Judge
    
    Call SaveLog("SKChange", "New Grade : " & strNew_Judge)
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("SKChange", ErrMsg)
    
    SKChange = strNew_Judge
    
    dbMyDB.Close
    
End Function

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Count change grade
'
'==========================================================================================================
Public Function Count_Change(ByVal pGrade As String) As String

    Dim typCOUNT_CHANGE             As COUNT_CHANGE_DATA
    
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strPath = App.PATH & "\Env\"
    strFileName = "Auto_Grade.cfg"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                With typCOUNT_CHANGE
                    Select Case Left(strTemp, intPos - 1)
                    Case "FINAL RANK":
                        .FINAL_GRADE = Mid(strTemp, intPos + 1)
                    Case "CHANGE GRADE":
                        .NEW_GRADE = Mid(strTemp, intPos + 1)
                    Case "COUNT":
                        .COUNT = CInt(Mid(strTemp, intPos + 1))
                    Case "CURRENT COUNT":
                        .CURRENT_COUNT = CInt(Mid(strTemp, intPos + 1))
                    End Select
                End With
            End If
        Wend
        
        Close intFileNum
        
        If typCOUNT_CHANGE.COUNT > 0 Then
            If pGrade = typCOUNT_CHANGE.FINAL_GRADE Then
                typCOUNT_CHANGE.CURRENT_COUNT = typCOUNT_CHANGE.CURRENT_COUNT + 1
                Count_Change = pGrade
                Call SaveLog("Count_Change", "Grade change from " & pGrade & " to " & typCOUNT_CHANGE.NEW_GRADE)
                Call SaveLog("            ", "Current change grade count : " & typCOUNT_CHANGE.CURRENT_COUNT)
                If typCOUNT_CHANGE.CURRENT_COUNT = typCOUNT_CHANGE.COUNT Then
                    typCOUNT_CHANGE.COUNT = 0
                End If
                intFileNum = FreeFile
                Open strPath & strFileName For Output As intFileNum
                
                With typCOUNT_CHANGE
                    strTemp = "FINAL RANK=" & .FINAL_GRADE
                    Print #intFileNum, strTemp
                    
                    strTemp = "CHANGE GRADE=" & .NEW_GRADE
                    Print #intFileNum, strTemp
                    
                    strTemp = "COUNT=" & .COUNT
                    Print #intFileNum, strTemp
                    
                    strTemp = "CURRENT COUNT=" & .CURRENT_COUNT
                    Print #intFileNum, strTemp
                    frmMain.StatusBar.Panels(2).Text = "Remained Grade Count : " & .CURRENT_COUNT
                    
                    Close intFileNum
                End With
            Else
                Count_Change = pGrade
            End If
        Else
            Count_Change = pGrade
        End If
    Else
        Count_Change = pGrade
    End If
    
    Exit Function
    
ErrorHandler:

    Count_Change = pGrade
    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Count_Change", ErrMsg)

End Function

Private Sub Update_SK_Flag()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strPANEL_DB_Path            As String
    Dim strPANEL_DB_FileName        As String
    Dim strQuery                    As String
    Dim strKEYID                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strKEYID = RANK_OBJ.Get_Current_KEYID
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "KEYID = '" & strKEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            strPANEL_DB_Path = lstRecord.Fields("PATH")
            strPANEL_DB_FileName = lstRecord.Fields("FILENAME")
            
            If Right(strPANEL_DB_Path, 1) <> "\" Then
                strPANEL_DB_Path = strPANEL_DB_Path & "\"
            End If
        End If
        lstRecord.Close
        
        dbMyDB.Close
        
        If Dir(strPANEL_DB_Path & strPANEL_DB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strPANEL_DB_Path & strPANEL_DB_FileName)
            
            strQuery = "UPDATE JOB_DATA SET "
            strQuery = strQuery & "SK_FLAG = '" & pubJOB_INFO.SK_FLAG & "'"
            
            dbMyDB.Execute (strQuery)
            
            dbMyDB.Close
        End If
    End If
    
End Sub
