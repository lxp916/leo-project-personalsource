Attribute VB_Name = "modDefect_File"
Option Explicit

Private Function FTP_Upload(ByVal pRemotePath As String, ByVal pLocalPath As String, ByVal pFileName As String) As Boolean

    Dim FTP_OBJ                 As New clsFTP
    
    Dim fe_object          As New clsFileExchanger
    
    Dim bolResult               As Boolean
     Dim isUpload               As Boolean
    
    Dim ErrMsg                  As String
    
On Error GoTo ErrorHandler

' Add upload mode selector, by leo.2012.08.15  'True:use FTP uploading mode, false:use common uploading mode
    If (fe_object.IsFTPUploadMode) Then
        If FTP_OBJ.Init_FTP_Client = True Then       'FTP Object Initialize
            If Right(pRemotePath, 1) <> "\" Then
                pRemotePath = pRemotePath & "\"
            End If
            Call FTP_OBJ.Open_Session                     'FTP Session Open
            bolResult = FTP_OBJ.FTP_Put_File(pFileName, pRemotePath, pLocalPath)
            If bolResult = False Then
                Call SaveLog("Defect_File_Upload", pFileName & " upload fail. Remote path : " & pRemotePath)
                FTP_Upload = False
            Else
                FTP_Upload = True
            End If
            FTP_OBJ.Close_Session
            FTP_OBJ.Disconnect_FTP_Client
        Else
            FTP_Upload = False
        End If
    Else ' common lan network uploadming mode added by leo 2012.08.19
        If Right(pRemotePath, 1) <> "\" Then
            pRemotePath = pRemotePath & "\"
        End If
        If fe_object.Check_Network = False Then
            Call Show_Message("Network is disconnected", "Remote server is not reachable, please check your network.")
            Call fe_object.Write_Local_Index(pFileName, pLocalPath, pRemotePath)
        Else
            Call fe_object.Write_Remote_Index(pFileName, pRemotePath)
            bolResult = fe_object.do_Upload(pFileName, pLocalPath, pRemotePath)
            
            If bolResult = False Then
                Call SaveLog("Defect_File_Upload", pFileName & " upload fail. Remote path : " & pRemotePath)
                Call fe_object.Write_Local_Index(pFileName, pLocalPath, pRemotePath)
                FTP_Upload = False
            Else

                If MsgBox("Do you want to upload local existing offline files?", vbYesNo, "") = vbYes Then
                   ' upload last faild files
                   If fe_object.hasExistingLocalIndex = True Then
                    Call fe_object.do_upload_files
                   End If
                End If
                Call SaveLog("Defect_File_Upload", pFileName & " upload successful. Remote path : " & pRemotePath)
                FTP_Upload = True
            End If

            
        End If
    End If
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("FTP_Upload", ErrMsg)
    FTP_Upload = False
    
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

Public Sub Make_Defect_File()

    Dim dbMyDB                          As Database
    
    Dim lstRecord                       As Recordset
    Dim lstPTN_Record                   As Recordset
    
    Dim typPFCD_DATA                    As PFCD_DATA
    
    Dim typPTN_DATA()                   As PATTERN_LIST_DATA
    Dim typDEFECT_INFO()                As DEFECT_DATA_STRUCTURE
    
    Dim typHEADER_DATA                  As DEFECT_FILE_HEADER
    Dim typPANEL_DATA                   As DEFECT_FILE_PANEL_DATA
    Dim typEQP_PANEL_DATA               As DEFECT_FILE_EQP_PANEL_DATA
    Dim typPANEL_SUMMARY                As DEFECT_FILE_PANEL_SUMMARY
    Dim typDEFECT_DATA()                As DEFECT_FILE_DEFECT_DATA
    Dim typLCD_DATA                     As DEFECT_FILE_LCD_DATA
    Dim typPANEL_PDS_SUMMARY            As DEFECT_FILE_PANEL_PDS_SUMMARY
    Dim typMAIN_DEFECT_DATA             As RANK_DATA_STRUCTURE
    
    Dim strDB_Path                      As String
    Dim strDB_FileName                  As String
    Dim strFilePath                     As String
    Dim strFileName                     As String
    Dim strQuery                        As String
    Dim strRUN_DATE                     As String
    Dim strRUN_TIME                     As String
    Dim strPath                         As String
    Dim strRemote_Path                  As String
    Dim strInspection_Time              As String
    Dim strPattern_Name                 As String
    Dim strMode_State                   As String
    Dim strMain_DEFECTCODE              As String
    Dim strMain_Rank                    As String
    Dim strPanel_Grade                  As String
    
    Dim intArray_Count                  As Integer
    Dim intPTN_Count                    As Integer
    Dim intIndex                        As Integer
    Dim intSubIndex                     As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
  
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PANEL_DATA WHERE "
        strQuery = strQuery & "KEYID = '" & RANK_OBJ.Get_Current_KEYID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            
            strMain_Rank = Trim(lstRecord.Fields("PANEL_RANK"))
            strMain_DEFECTCODE = Trim(lstRecord.Fields("PANEL_LOSSCODE"))
            strPanel_Grade = Trim(lstRecord.Fields("PANEL_GRADE"))
            strRUN_DATE = Trim(Str$(lstRecord.Fields("RUN_DATE")))
            strRUN_TIME = Trim(Str$(lstRecord.Fields("RUN_TIME")))
            strFilePath = lstRecord.Fields("PATH")
            strFileName = lstRecord.Fields("FILENAME")
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    With typHEADER_DATA
        .JPS_VERSION = App.Major & "." & App.Minor & "." & App.Revision
        .FILE_CREATE_TIME = Format(DATE, "YYYY/MM/DD") & "_" & Format(TIME, "HH:MM:SS")
        .EQUIP_TYPE = Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5)
        .EQ_ID = frmMain.flxEQ_Information.TextMatrix(5, 1)
        .SUBEQ_ID = .EQ_ID
    End With
    With typPANEL_DATA
        .PANELID = Trim(pubPANEL_INFO.PANELID)
        .GLASS_TYPE = Trim(pubSHARE_INFO.GLASS_TYPE)
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
    With typEQP_PANEL_DATA
        .RECIPE_NO = Trim(pubSHARE_INFO.RECIPEID)
        .START_TIME = Trim(RANK_OBJ.Get_START_TIME)
        .END_TIME = Trim(RANK_OBJ.Get_END_TIME)
        .OPERATOR_ID = Trim(frmMain.lblUser.Caption)
        .TOTAL_POINT_DEFECT_COUNT = Trim(RANK_OBJ.Get_TB_Count + RANK_OBJ.Get_TD_Count)
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
        .OPERATION_MODE = strMode_State
        .MAIN_DEFECT_CODE = strMain_DEFECTCODE
        .RJS_NO = ""
        .INSPECTION_TIME = Trim(Format(RANK_OBJ.Get_Tact_Time, "0.00"))
        .TRANSFER_TIME = Trim(Format(RANK_OBJ.Get_Transfer_Time, "0.00"))
        .TACT_TIME = Trim(Format(CDbl(.INSPECTION_TIME) + CDbl(.TRANSFER_TIME), "0.00"))
    End With
    Call Get_PFCD_DATA(typPFCD_DATA, pubCST_INFO.PFCD)
    With typPANEL_SUMMARY
        .PANELID = Trim(pubPANEL_INFO.PANELID)
       'for CAAPI Judge Rank modify----Lucas.2011.11.18
        If strMode_State = "FA" Then
            .JUDGE_RANK = "Y "
        Else
            .JUDGE_RANK = strMain_Rank
        End If
       'for CAAPI Judge Rank modify----Lucas.2011.11.18
       
        .MAIN_DEFECT_CODE = strMain_DEFECTCODE
        .DATA_TOTAL_PIXEL = Trim(typPFCD_DATA.X_PIXEL_LENGTH)
        .GATE_TOTAL_PIXEL = Trim(typPFCD_DATA.Y_PIXEL_LENGTH)
        .DRIVE_TYPE = Trim(frmMain.flxEQ_Information.TextMatrix(1, 1))
        .BACKLIGHT = Trim(EQP.Get_BackLight_Value)
'        .LIGHT_ON_TARGET_REASON_CODE = Trim(.MAIN_DEFECT_CODE)
'Lucas Ver.0.9.34 2012.06.26---------For Light ON PreGrade
         .LIGHT_ON_PRE_GRADE = Trim(pubJOB_INFO.JOB_GRADE)
'Lucas Ver.0.9.34 2012.06.26---------For Light ON PreGrade
        typMAIN_DEFECT_DATA = Get_DEFECT_DATA_by_CODE(.MAIN_DEFECT_CODE)
        .LIGHT_ON_TARGET_REASON_TYPE = Trim(typMAIN_DEFECT_DATA.DEFECT_NAME)
        .SLOT_ID = Trim(pubPANEL_INFO.SLOT_NUM)
    End With
    
    intArray_Count = 0
    If Dir(strFilePath & strFileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strFilePath & strFileName)
    
        strQuery = "SELECT * FROM DEFECT_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            intArray_Count = lstRecord.RecordCount
            ReDim typDEFECT_INFO(intArray_Count)
            ReDim typDEFECT_DATA(intArray_Count)
            
            lstRecord.Close
            
            intPTN_Count = EQP.Get_PATTERN_COUNT
            For intIndex = 1 To intArray_Count
                strQuery = "SELECT * FROM DEFECT_DATA WHERE "
                strQuery = strQuery & "DEFECT_NO=" & intIndex
                
                Set lstRecord = dbMyDB.OpenRecordset(strQuery)
                
                If lstRecord.EOF = False Then
                    lstRecord.MoveFirst
                    With lstRecord
                        If (.Fields("DEFECT_CODE") <> "CDBDT") And (.Fields("DEFECT_CODE") <> "CDDKD") And (.Fields("DEFECT_CODE") <> "CDBDD") And _
                           (.Fields("DEFECT_CODE") <> "CDBTT") And (.Fields("DEFECT_CODE") <> "CDDKT") And (.Fields("DEFECT_CODE") <> "CDBDT") Then
                            typDEFECT_DATA(intIndex).PANELID = Trim(.Fields("PANELID"))
                            typDEFECT_DATA(intIndex).DEFECT_NO = Trim(.Fields("DEFECT_NO"))
                            typDEFECT_DATA(intIndex).DEFECT_CODE = Trim(.Fields("DEFECT_CODE"))
                            typDEFECT_DATA(intIndex).COLOR = Trim(.Fields("COLOR"))
                            typDEFECT_DATA(intIndex).PANEL_GRADE = Trim(.Fields("DEFECT_GRADE"))
                            typDEFECT_DATA(intIndex).GRAY_LEVEL = Trim(.Fields("GRAY_LEVEL"))
                            typDEFECT_DATA(intIndex).DEFECT_NAME = Trim(.Fields("DEFECT_NAME"))
                            typDEFECT_DATA(intIndex).DATA_X1 = Trim(.Fields("DEFECT_DATA1"))
                            typDEFECT_DATA(intIndex).DATA_X2 = Trim(.Fields("DEFECT_DATA2"))
                            typDEFECT_DATA(intIndex).DATA_X3 = Trim(.Fields("DEFECT_DATA3"))
                            typDEFECT_DATA(intIndex).GATE_Y1 = Trim(.Fields("DEFECT_GATE1"))
                            typDEFECT_DATA(intIndex).GATE_Y2 = Trim(.Fields("DEFECT_GATE2"))
                            typDEFECT_DATA(intIndex).GATE_Y3 = Trim(.Fields("DEFECT_GATE3"))
                            Select Case Mid(typDEFECT_DATA(intIndex).DEFECT_CODE, 2, 1)
                            Case "D":
                                typDEFECT_DATA(intIndex).MARK_TYPE = "Point"
                            Case "L":
                                typDEFECT_DATA(intIndex).MARK_TYPE = "Line"
                            Case "M":
                                typDEFECT_DATA(intIndex).MARK_TYPE = "Rect"
                            Case "G":
                                typDEFECT_DATA(intIndex).MARK_TYPE = "Circle"
                            Case Else
                                typDEFECT_DATA(intIndex).MARK_TYPE = "Point"
                            End Select
                            
                            typDEFECT_INFO(intIndex).DEFECT_CODE = typDEFECT_DATA(intIndex).DEFECT_CODE
                            
                        End If
                    End With
                    strQuery = "SELECT * FROM PATTERN_INSPECTION"
                    
                    Set lstPTN_Record = dbMyDB.OpenRecordset(strQuery)
                    
                    If lstPTN_Record.EOF = False Then
                        lstPTN_Record.MoveLast
                        intPTN_Count = lstPTN_Record.RecordCount
                        ReDim typDEFECT_DATA(intIndex).EACH_PTN_INSPECTION_TIME(intPTN_Count)
                        ReDim typDEFECT_DATA(intIndex).PATTERN_NAME(intPTN_Count)
                        
                        lstPTN_Record.MoveFirst
                        intSubIndex = 0
                        While lstPTN_Record.EOF = False
                            intSubIndex = intSubIndex + 1
                            typDEFECT_DATA(intIndex).EACH_PTN_INSPECTION_TIME(intSubIndex) = lstPTN_Record.Fields("INSPECTION_TIME")
                            typDEFECT_DATA(intIndex).PATTERN_NAME(intSubIndex) = lstPTN_Record.Fields("PATTERN_NAME")
                            lstPTN_Record.MoveNext
                        Wend
                    End If
                    lstPTN_Record.Close
                End If
                lstRecord.Close
            Next intIndex
        Else
            lstRecord.Close
        End If
        
        dbMyDB.Close
    End If
    
    With typLCD_DATA
        .PANELID = Trim(pubPANEL_INFO.PANELID)
        .LIGHT_ON_SOURCE_GRADE = strPanel_Grade
        .LIGHT_ON_SOURCE_REASON_CODE = strMain_DEFECTCODE
        .TOTAL_LIGHT_ON_DEFECT_COUNT = intArray_Count
    End With
    With typPANEL_PDS_SUMMARY
        .PANELID = Trim(pubPANEL_INFO.PANELID)
        .PARAMETER_NAME = Trim(pubCST_INFO.OWNER)
        .AVG = Trim(pubCST_INFO.CST_SPARE(1))
        .MIN = Trim(pubCST_INFO.CST_SPARE(2))
        .MAX = Trim(pubCST_INFO.CST_SPARE(3))
        .STD = Trim(pubCST_INFO.CST_SPARE(4))
        .Count = Trim(pubPANEL_INFO.REPAIR_REWORK_COUNT)
    End With
        
    If strRUN_DATE = "" Then
        strRUN_DATE = Format(DATE, "YYYYMMDD")
    End If
    If strRUN_TIME = "" Then
        strRUN_TIME = Format(TIME, "HHMMSS")
    End If
    
    If typHEADER_DATA.EQUIP_TYPE = "CATST" Then
        strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & pubPANEL_INFO.PRODUCTID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Source\"
        strPath = App.PATH & "\DB\" & CInt(Mid(strRUN_DATE, 5, 2)) & "\" & CInt(Mid(strRUN_DATE, 7, 2)) & "\"
        strFileName = pubCST_INFO.PROCESS_NUM & "_" & pubPANEL_INFO.PANELID & "_" & strRUN_DATE & "_" & strRUN_TIME & ".csv"
        
        Call Write_Header_CATST(strPath, strFileName, typHEADER_DATA)
        Call Write_Panel_Data_CATST(strPath, strFileName, typPANEL_DATA)
        Call Write_EQP_Panel_Data_CATST(strPath, strFileName, typEQP_PANEL_DATA)
        Call Write_Panel_Summary_CATST(strPath, strFileName, typPANEL_SUMMARY)
        Call Write_Defect_Data_CATST(strPath, strFileName, intArray_Count, typDEFECT_DATA, typPANEL_DATA.PRODUCT_ID, pubPANEL_INFO.PANELID)
        Call Write_LCD_Data_CATST(strPath, strFileName, typLCD_DATA)
        Call Write_Panel_PDS_Summay_CATST(strPath, strFileName, typPANEL_PDS_SUMMARY)
        
        If EQP.Get_DEFECT_UPLOAD = True Then
              If FTP_Upload(strRemote_Path, strPath, strFileName) = True Then
                Call SaveLog("Make_Defect_File", strFileName & " upload complete.")
              Else
                Call SaveLog("Make_Defect_File", strFileName & " upload fail.")
              End If
         End If
         If strMode_State = "FA" Then
               strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
               strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Source\"
               If Make_FTP_Folder(strRemote_Path) = True Then
               Call SaveLog("Make_Defect_File", strRemote_Path & " folder create complete.")
               Else
               Call SaveLog("Make_Defect_File", strRemote_Path & " folder create fail.")
               End If
         End If
        
        'Lucas 2011.12.26  Ver.0.7.33---Path change for CANRP/CALOI
        'Path and FileName change
'        strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
'        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Share\"

        
        strRemote_Path = "Link\" & "CATST\" & Mid(typPANEL_DATA.PRODUCT_ID, 3, 5) & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\"
        strFileName = pubPANEL_INFO.PANELID & ".csv"
        
        Call Write_Share_Header_CATST(strPath, strFileName, typHEADER_DATA)
        Call Write_Share_Panel_Summary_CATST(strPath, strFileName, typPANEL_SUMMARY)
        Call Write_Share_Defect_Data_CATST(strPath, strFileName, intArray_Count, typDEFECT_DATA, typPANEL_DATA.PRODUCT_ID, pubPANEL_INFO.PANELID)
        Call Write_LCD_Data_CATST(strPath, strFileName, typLCD_DATA)
        Call Write_Panel_PDS_Summay_CATST(strPath, strFileName, typPANEL_PDS_SUMMARY)
    
        If EQP.Get_DEFECT_UPLOAD = True Then
            If FTP_Upload(strRemote_Path, strPath, strFileName) = True Then
                Call SaveLog("Make_Defect_File", strFileName & " upload complete.")
            Else
                Call SaveLog("Make_Defect_File", strFileName & " upload fail.")
            End If
        End If
    Else
        strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Source\"
        strPath = App.PATH & "\DB\" & CInt(Mid(strRUN_DATE, 5, 2)) & "\" & CInt(Mid(strRUN_DATE, 7, 2)) & "\"
        strFileName = typPANEL_DATA.PROCESS_ID & "_" & pubPANEL_INFO.PANELID & "_" & strRUN_DATE & "_" & strRUN_TIME & ".csv"
        
        Call Write_Header_CALOI(strPath, strFileName, typHEADER_DATA)
        Call Write_Panel_Data_CALOI(strPath, strFileName, typPANEL_DATA)
        Call Write_EQP_Panel_Data_CALOI(strPath, strFileName, typEQP_PANEL_DATA)
        Call Write_Panel_Summary_CALOI(strPath, strFileName, typPANEL_SUMMARY)
        Call Write_Defect_Data_CALOI(strPath, strFileName, intArray_Count, typDEFECT_DATA, typPANEL_DATA.PRODUCT_ID, pubPANEL_INFO.PANELID)
        Call Write_LCD_Data_CALOI(strPath, strFileName, typLCD_DATA)
        Call Write_Panel_PDS_Summay_CALOI(strPath, strFileName, typPANEL_PDS_SUMMARY)
        
        If FTP_Upload(strRemote_Path, strPath, strFileName) = True Then
            Call SaveLog("Make_Defect_File", strFileName & " upload complete.")
        Else
            Call SaveLog("Make_Defect_File", strFileName & " upload fail.")
        End If
        
'       Lucas 2011.12.26  Ver.0.7.33---Path change for CANRP/CALOI
'       Path and FileName change
'        strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
'        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Share\"

        strRemote_Path = "Link\" & "CATST\" & Mid(typPANEL_DATA.PRODUCT_ID, 3, 5) & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
        strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\"
        strFileName = pubPANEL_INFO.PANELID & ".csv"
        
        Call Write_Share_Header_CALOI(strPath, strFileName, typHEADER_DATA)
        Call Write_Share_Panel_Summary_CALOI(strPath, strFileName, typPANEL_SUMMARY)
        Call Write_Share_Defect_Data_CALOI(strPath, strFileName, intArray_Count, typDEFECT_DATA, typPANEL_DATA.PRODUCT_ID, pubPANEL_INFO.PANELID)
        Call Write_LCD_Data_CALOI(strPath, strFileName, typLCD_DATA)
        Call Write_Panel_PDS_Summay_CALOI(strPath, strFileName, typPANEL_PDS_SUMMARY)
    
        If FTP_Upload(strRemote_Path, strPath, strFileName) = True Then
            Call SaveLog("Make_Defect_File", strFileName & " upload complete.")
        Else
            Call SaveLog("Make_Defect_File", strFileName & " upload fail.")
        End If
    End If
    
    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Image\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create fail.")
    End If
    
    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Error\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create fail.")
    End If
    
    strRemote_Path = typHEADER_DATA.EQUIP_TYPE & "\" & typPANEL_DATA.PRODUCT_ID & "\" & Mid(pubPANEL_INFO.PANELID, 1, 5) & "\"
    strRemote_Path = strRemote_Path & Mid(pubPANEL_INFO.PANELID, 1, 8) & "\" & pubPANEL_INFO.PANELID & "\Backup\"
    If Make_FTP_Folder(strRemote_Path) = True Then
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create complete.")
    Else
        Call SaveLog("Make_Defect_File", strRemote_Path & " folder create fail.")
    End If

    Call Check_Auto_Alarm       '2012.03.26 Added by K.H.KIM
    
End Sub

Private Sub Write_Share_Header_CATST(ByVal pPath As String, ByVal pFileName As String, pHEADER_DATA As DEFECT_FILE_HEADER)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Output As intFileNum
    
    strTemp = "HEADER_BEGIN"
    Print #intFileNum, strTemp
    
    With pHEADER_DATA
        strTemp = "FILE_VERSION:" & .JPS_VERSION
        Print #intFileNum, strTemp
        
        strTemp = "FILE_CREATED_TIME:" & .FILE_CREATE_TIME
        Print #intFileNum, strTemp
        
        strTemp = "EQUIP_TYPE:" & .EQUIP_TYPE
        Print #intFileNum, strTemp
        
        strTemp = "EQ_ID:" & .EQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "SUBEQ_ID:" & .SUBEQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "CONTENT:PANEL_SUMMARY/DEFECT_DATA/LCD_DATA/PANEL_PDS_SUMMARY"
        Print #intFileNum, strTemp
    End With
    
    strTemp = "HEADER_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Share_Panel_Summary_CATST(ByVal pPath As String, ByVal pFileName As String, pPANEL_SUMMARY As DEFECT_FILE_PANEL_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,JUDGE_RANK,MAIN_DEF_CODE,DATA_TOTAL_PIXEL,GATE_TOTAL_PIXEL,DRIVE_TYPE,LIGHT_ON_PRE_GRADE,LIGHT_ON_TARGET_REASON_TYPE_E"
    Print #intFileNum, strTemp
    
    With pPANEL_SUMMARY
        strTemp = .PANELID & "," & .JUDGE_RANK & "," & .MAIN_DEFECT_CODE & "," & .DATA_TOTAL_PIXEL & "," & .GATE_TOTAL_PIXEL & "," & .DRIVE_TYPE & ","
        strTemp = strTemp & .LIGHT_ON_PRE_GRADE & "," & .LIGHT_ON_TARGET_REASON_TYPE
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_SUMMARY_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Share_Defect_Data_CATST(ByVal pPath As String, ByVal pFileName As String, ByVal pArray_Count As Integer, pDEFECT_DATA() As DEFECT_FILE_DEFECT_DATA, _
                                          ByVal pPRODUCTID As String, ByVal pPanelID As String)

    Dim PFCD_ADDRESS_DATA               As PFCD_ADDRESS_STRUCTURE
    
    Dim intFileNum                      As Integer
    Dim intIndex                        As Integer
    Dim intPTN_Index                    As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "DEFECT_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,DEFECT_NO,DEFECT_CODE,COLOR,PANEL_GRADE,GRAY_LEVEL,DEFECT_NAME,DATA_X1,GATE_Y1,DATA_X2,GATE_Y2,DATA_X3,GATE_Y3,"
    strTemp = strTemp & "PANEL_COORDINATE_X1,PANEL_COORDINATE_Y1,PANEL_COORDINATE_X2,PANEL_COORDINDATE_Y2,PANEL_COORDINATE_X3,PANEL_COORDINATE_Y3,"
    strTemp = strTemp & "GLASS_COORDINATE_X1,GLASS_COORDINATE_Y1,GLASS_COORDINATE_X2,GLASS_COORDINDATE_Y2,GLASS_COORDINATE_X3,GLASS_COORDINATE_Y3,"
    strTemp = strTemp & "MARK_TYPE,EACH_PATTERN_INSPECTION_TIME,PATTERN_NAME"
    Print #intFileNum, strTemp
    
    Call Get_PFCD_ADDRESS_DATA(PFCD_ADDRESS_DATA, pPRODUCTID, CInt(Right(pPanelID, 2)))
    For intIndex = 1 To pArray_Count
        With pDEFECT_DATA(intIndex)
            If .DEFECT_CODE <> "" Then
                If Trim(.DATA_X1) <> "" Then
                    .PANEL_COORDINATE_X1 = CDbl(.DATA_X1) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                End If
                If Trim(.GATE_Y1) <> "" Then
                    .PANEL_COORDINATE_Y1 = CDbl(.GATE_Y1) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                End If
                If Trim(.DATA_X2) <> "" Then
                    .PANEL_COORDINATE_X2 = CDbl(.DATA_X2) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                End If
                If Trim(.GATE_Y2) <> "" Then
                    .PANEL_COORDINATE_Y2 = CDbl(.GATE_Y2) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                End If
                If Trim(.DATA_X3) <> "" Then
                    .PANEL_COORDINATE_X3 = CDbl(.DATA_X3) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                End If
                If Trim(.GATE_Y3) <> "" Then
                    .PANEL_COORDINATE_Y3 = CDbl(.GATE_Y3) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                End If
                
                strTemp = .PANELID & "," & .DEFECT_NO & "," & .DEFECT_CODE & "," & .COLOR & "," & .PANEL_GRADE & "," & .GRAY_LEVEL & "," & .DEFECT_NAME & ","
                strTemp = strTemp & .DATA_X1 & "," & .GATE_Y1 & "," & .DATA_X2 & "," & .GATE_Y2 & "," & .DATA_X3 & "," & .GATE_Y3 & ","
                strTemp = strTemp & .PANEL_COORDINATE_X1 & "," & .PANEL_COORDINATE_Y1 & "," & .PANEL_COORDINATE_X2 & "," & .PANEL_COORDINATE_Y2 & "," & .PANEL_COORDINATE_X3 & "," & .PANEL_COORDINATE_Y3 & ","
                strTemp = strTemp & .GLASS_COORDINATE_X1 & "," & .GLASS_COORDINATE_Y1 & "," & .GLASS_COORDINATE_X2 & "," & .GLASS_COORDINATE_Y2 & "," & .GLASS_COORDINATE_X3 & "," & .GLASS_COORDINATE_Y3 & ","
                strTemp = strTemp & .MARK_TYPE & ","
                If .PATTERN_COUNT > 0 Then
                    For intPTN_Index = 1 To .PATTERN_COUNT
                        strTemp = strTemp & .EACH_PTN_INSPECTION_TIME(intPTN_Index) & ","
                    Next intPTN_Index
                    For intPTN_Index = 1 To .PATTERN_COUNT - 1
                        strTemp = strTemp & .PATTERN_NAME(intPTN_Index) & ","
                    Next intPTN_Index
                    strTemp = strTemp & .PATTERN_NAME(.PATTERN_COUNT)
                Else
                    strTemp = strTemp & ","
                End If
                Print #intFileNum, strTemp
            End If
        End With
    Next intIndex
    
    strTemp = "DEFECT_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Share_Header_CALOI(ByVal pPath As String, ByVal pFileName As String, pHEADER_DATA As DEFECT_FILE_HEADER)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "HEADER_BEGIN"
    Print #intFileNum, strTemp
    
    With pHEADER_DATA
        strTemp = "FILE_VERSION:" & .JPS_VERSION
        Print #intFileNum, strTemp
        
        strTemp = "FILE_CREATED_TIME:" & .FILE_CREATE_TIME
        Print #intFileNum, strTemp
        
        strTemp = "EQUIP_TYPE:" & .EQUIP_TYPE
        Print #intFileNum, strTemp
        
        strTemp = "EQ_ID:" & .EQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "SUBEQ_ID:" & .SUBEQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "CONTENT:PANEL_SUMMARY/DEFECT_DATA/LCD_DATA/PANEL_PDS_SUMMARY"
        Print #intFileNum, strTemp
    End With
    
    strTemp = "HEADER_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Share_Panel_Summary_CALOI(ByVal pPath As String, ByVal pFileName As String, pPANEL_SUMMARY As DEFECT_FILE_PANEL_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,JUDGE_RANK,MAIN_DEF_CODE,DATA_TOTAL_PIXEL,GATE_TOTAL_PIXEL,DRIVE_TYPE,LIGHT_ON_PRE_GRADE,LIGHT_ON_TARGET_REASON_TYPE_E"
    Print #intFileNum, strTemp
    
    With pPANEL_SUMMARY
        strTemp = .PANELID & "," & .JUDGE_RANK & "," & .MAIN_DEFECT_CODE & "," & .DATA_TOTAL_PIXEL & "," & .GATE_TOTAL_PIXEL & "," & .DRIVE_TYPE & ","
        strTemp = strTemp & .LIGHT_ON_PRE_GRADE & "," & .LIGHT_ON_TARGET_REASON_TYPE
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_SUMMARY_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Share_Defect_Data_CALOI(ByVal pPath As String, ByVal pFileName As String, ByVal pArray_Count As Integer, pDEFECT_DATA() As DEFECT_FILE_DEFECT_DATA, _
                                          ByVal pPRODUCTID As String, ByVal pPanelID As String)

    Dim PFCD_ADDRESS_DATA               As PFCD_ADDRESS_STRUCTURE
    
    Dim intFileNum                      As Integer
    Dim intIndex                        As Integer
    Dim intPTN_Index                    As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "DEFECT_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,DEFECT_NO,DEFECT_CODE,COLOR,PANEL_GRADE,GRAY_LEVEL,DEFECT_NAME,DATA_X1,GATE_Y1,DATA_X2,GATE_Y2,DATA_X3,GATE_Y3,"
    strTemp = strTemp & "PANEL_COORDINATE_X1,PANEL_COORDINATE_Y1,PANEL_COORDINATE_X2,PANEL_COORDINDATE_Y2,PANEL_COORDINATE_X3,PANEL_COORDINATE_Y3,"
    strTemp = strTemp & "GLASS_COORDINATE_X1,GLASS_COORDINATE_Y1,GLASS_COORDINATE_X2,GLASS_COORDINDATE_Y2,GLASS_COORDINATE_X3,GLASS_COORDINATE_Y3,"
    strTemp = strTemp & "MARK_TYPE,EACH_PATTERN_INSPECTION_TIME,PATTERN_NAME"
    Print #intFileNum, strTemp
    
    Call Get_PFCD_ADDRESS_DATA(PFCD_ADDRESS_DATA, pPRODUCTID, CInt(Right(pPanelID, 2)))
    For intIndex = 1 To pArray_Count
        With pDEFECT_DATA(intIndex)
            If .DEFECT_CODE <> "" Then
                If Trim(.DATA_X1) <> "" Then
                    .PANEL_COORDINATE_X1 = CDbl(.DATA_X1) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                End If
                If Trim(.GATE_Y1) <> "" Then
                    .PANEL_COORDINATE_Y1 = CDbl(.GATE_Y1) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                End If
                If Trim(.DATA_X2) <> "" Then
                    .PANEL_COORDINATE_X2 = CDbl(.DATA_X2) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                End If
                If Trim(.GATE_Y2) <> "" Then
                    .PANEL_COORDINATE_Y2 = CDbl(.GATE_Y2) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                End If
                If Trim(.DATA_X3) <> "" Then
                    .PANEL_COORDINATE_X3 = CDbl(.DATA_X3) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                    .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                End If
                If Trim(.GATE_Y3) <> "" Then
                    .PANEL_COORDINATE_Y3 = CDbl(.GATE_Y3) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                    .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                End If
                
                strTemp = .PANELID & "," & .DEFECT_NO & "," & .DEFECT_CODE & "," & .COLOR & "," & .PANEL_GRADE & "," & .GRAY_LEVEL & "," & .DEFECT_NAME & ","
                strTemp = strTemp & .DATA_X1 & "," & .GATE_Y1 & "," & .DATA_X2 & "," & .GATE_Y2 & "," & .DATA_X3 & "," & .GATE_Y3 & ","
                strTemp = strTemp & .PANEL_COORDINATE_X1 & "," & .PANEL_COORDINATE_Y1 & "," & .PANEL_COORDINATE_X2 & "," & .PANEL_COORDINATE_Y2 & "," & .PANEL_COORDINATE_X3 & "," & .PANEL_COORDINATE_Y3 & ","
                strTemp = strTemp & .GLASS_COORDINATE_X1 & "," & .GLASS_COORDINATE_Y1 & "," & .GLASS_COORDINATE_X2 & "," & .GLASS_COORDINATE_Y2 & "," & .GLASS_COORDINATE_X3 & "," & .GLASS_COORDINATE_Y3 & ","
                strTemp = strTemp & .MARK_TYPE & ","
                If .PATTERN_COUNT > 0 Then
                    For intPTN_Index = 1 To .PATTERN_COUNT
                        strTemp = strTemp & .EACH_PTN_INSPECTION_TIME(intPTN_Index) & ","
                    Next intPTN_Index
                    For intPTN_Index = 1 To .PATTERN_COUNT - 1
                        strTemp = strTemp & .PATTERN_NAME(intPTN_Index) & ","
                    Next intPTN_Index
                    strTemp = strTemp & .PATTERN_NAME(.PATTERN_COUNT)
                Else
                    strTemp = strTemp & ","
                End If
                Print #intFileNum, strTemp
            End If
        End With
    Next intIndex
    
    strTemp = "DEFECT_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Header_CATST(ByVal pPath As String, ByVal pFileName As String, pHEADER_DATA As DEFECT_FILE_HEADER)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Output As intFileNum
    
    strTemp = "HEADER_BEGIN"
    Print #intFileNum, strTemp
    
    With pHEADER_DATA
        strTemp = "FILE_VERSION:" & .JPS_VERSION
        Print #intFileNum, strTemp
        
        strTemp = "FILE_CREATED_TIME:" & .FILE_CREATE_TIME
        Print #intFileNum, strTemp
        
        strTemp = "EQUIP_TYPE:" & .EQUIP_TYPE
        Print #intFileNum, strTemp
        
        strTemp = "EQ_ID:" & .EQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "SUBEQ_ID:" & .SUBEQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "CONTENT:PANEL_DATA/EQP_PANEL_DATA/PANEL_SUMMARY/DEFECT_DATA/LCD_DATA/PANEL_PDS_SUMMARY"
        Print #intFileNum, strTemp
    End With
    
    strTemp = "HEADER_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum
    
End Sub

Private Sub Write_Panel_Data_CATST(ByVal pPath As String, ByVal pFileName As String, pPANEL_DATA As DEFECT_FILE_PANEL_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,GLASS_TYPE,PRODUCT_ID,PROCESS_ID,RECIPE_ID,SALEORDER,CF_GLASS_ID,ARRAY_LOT_ID,ARRAY_GLASS_ID,CF_GLASS_INFO,TFT_REPAIR_GRADE, GROUPID"
    Print #intFileNum, strTemp
        
    With pPANEL_DATA
        strTemp = .PANELID & "," & .GLASS_TYPE & "," & .PRODUCT_ID & "," & .PROCESS_ID & "," & .RECIPE_ID & "," & .SALEORDER & ","
        strTemp = strTemp & .CF_GLASS_ID & "," & .ARRAY_LOT_ID & "," & .ARRAY_GLASS_ID & "," & .CF_GLASS_OX_INFO & "," & .TFT_PANEL_JUDGE & "," & .GROUP_ID
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum
    
End Sub

Private Sub Write_EQP_Panel_Data_CATST(ByVal pPath As String, ByVal pFileName As String, pEQP_PANEL_DATA As DEFECT_FILE_EQP_PANEL_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "EQP_PANEL_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "RECIPE_NO,START_TIME,END_TIME,OPERATOR_ID,TACT_TIME,MAIN_DEFECT_CODE,TOTAL_POINT_DEF_CNT,OPERATION_MODE,RJS_NO,INSPECTION_TIME,TRANSFER_TIME"
    Print #intFileNum, strTemp
    
    With pEQP_PANEL_DATA
        strTemp = .RECIPE_NO & "," & .START_TIME & "," & .END_TIME & "," & .OPERATOR_ID & "," & .TACT_TIME & "," & .MAIN_DEFECT_CODE & ","
        strTemp = strTemp & .TOTAL_POINT_DEFECT_COUNT & "," & .OPERATION_MODE & "," & .RJS_NO & "," & .INSPECTION_TIME & "," & .TRANSFER_TIME
        Print #intFileNum, strTemp
    End With
    
    strTemp = "EQP_PANEL_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Panel_Summary_CATST(ByVal pPath As String, ByVal pFileName As String, pPANEL_SUMMARY As DEFECT_FILE_PANEL_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,JUDGE_RANK,MAIN_DEF_CODE,DATA_TOTAL_PIXEL,GATE_TOTAL_PIXEL,DRIVE_TYPE,BACKLIGHT,LIGHT_ON_PRE_GRADE,LIGHT_ON_TARGET_REASON_TYPE_E,SLOT_ID"
    Print #intFileNum, strTemp
    
    With pPANEL_SUMMARY
        strTemp = .PANELID & "," & .JUDGE_RANK & "," & .MAIN_DEFECT_CODE & "," & .DATA_TOTAL_PIXEL & "," & .GATE_TOTAL_PIXEL & "," & .DRIVE_TYPE & ","
        strTemp = strTemp & .BACKLIGHT & "," & .LIGHT_ON_PRE_GRADE & "," & .LIGHT_ON_TARGET_REASON_TYPE & "," & .SLOT_ID
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_SUMMARY_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Defect_Data_CATST(ByVal pPath As String, ByVal pFileName As String, ByVal pArray_Count As Integer, _
                                    pDEFECT_DATA() As DEFECT_FILE_DEFECT_DATA, ByVal pPRODUCTID As String, ByVal pPanelID As String)

    Dim PFCD_ADDRESS_DATA               As PFCD_ADDRESS_STRUCTURE
    
    Dim intFileNum                      As Integer
    Dim intIndex                        As Integer
    Dim intPTN_Index                    As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "DEFECT_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,DEFECT_NO,DEFECT_CODE,COLOR,PANEL_GRADE,GRAY_LEVEL,DEFECT_NAME,DATA_X1,GATE_Y1,DATA_X2,GATE_Y2,DATA_X3,GATE_Y3,"
    strTemp = strTemp & "PANEL_COORDINATE_X1,PANEL_COORDINATE_Y1,PANEL_COORDINATE_X2,PANEL_COORDINDATE_Y2,PANEL_COORDINATE_X3,PANEL_COORDINATE_Y3,"
    strTemp = strTemp & "GLASS_COORDINATE_X1,GLASS_COORDINATE_Y1,GLASS_COORDINATE_X2,GLASS_COORDINDATE_Y2,GLASS_COORDINATE_X3,GLASS_COORDINATE_Y3,"
    strTemp = strTemp & "MARK_TYPE,EACH_PATTERN_INSPECTION_TIME,PATTERN_NAME"
    Print #intFileNum, strTemp
    
    Call Get_PFCD_ADDRESS_DATA(PFCD_ADDRESS_DATA, pPRODUCTID, CInt(Right(pPanelID, 2)))
    For intIndex = 1 To pArray_Count
        With pDEFECT_DATA(intIndex)
            If .DEFECT_CODE <> "" Then
                If Trim(.DATA_X1) <> "" Then
                    .PANEL_COORDINATE_X1 = CDbl(.DATA_X1) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y1) <> "" Then
                    .PANEL_COORDINATE_Y1 = CDbl(.GATE_Y1) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If Trim(.DATA_X2) <> "" Then
                    .PANEL_COORDINATE_X2 = CDbl(.DATA_X2) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y2) <> "" Then
                    .PANEL_COORDINATE_Y2 = CDbl(.GATE_Y2) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If Trim(.DATA_X3) <> "" Then
                    .PANEL_COORDINATE_X3 = CDbl(.DATA_X3) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y3) <> "" Then
                    .PANEL_COORDINATE_Y3 = CDbl(.GATE_Y3) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If PFCD_ADDRESS_DATA.SOURCE_DIRECTION = "V" Then
                    Select Case PFCD_ADDRESS_DATA.ORIGIN_LOCATION
                    Case "LT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X3)
                        End If
                    Case "LB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                        End If
                    Case "RT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X3)
                        End If
                    Case "RB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                        End If
                    End Select
                Else
                    Select Case PFCD_ADDRESS_DATA.ORIGIN_LOCATION
                    Case "LT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y3)
                        End If
                    Case "LB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y3)
                        End If
                    Case "RT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y3)
                        End If
                    Case "RB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y3)
                        End If
                    End Select
                End If
                
                strTemp = .PANELID & "," & .DEFECT_NO & "," & .DEFECT_CODE & "," & .COLOR & "," & .PANEL_GRADE & "," & .GRAY_LEVEL & "," & .DEFECT_NAME & ","
                strTemp = strTemp & .DATA_X1 & "," & .GATE_Y1 & "," & .DATA_X2 & "," & .GATE_Y2 & "," & .DATA_X3 & "," & .GATE_Y3 & ","
                strTemp = strTemp & .PANEL_COORDINATE_X1 & "," & .PANEL_COORDINATE_Y1 & "," & .PANEL_COORDINATE_X2 & "," & .PANEL_COORDINATE_Y2 & "," & .PANEL_COORDINATE_X3 & "," & .PANEL_COORDINATE_Y3 & ","
                strTemp = strTemp & .GLASS_COORDINATE_X1 & "," & .GLASS_COORDINATE_Y1 & "," & .GLASS_COORDINATE_X2 & "," & .GLASS_COORDINATE_Y2 & "," & .GLASS_COORDINATE_X3 & "," & .GLASS_COORDINATE_Y3 & ","
                strTemp = strTemp & .MARK_TYPE & ","
                If .PATTERN_COUNT > 0 Then
                    For intPTN_Index = 1 To .PATTERN_COUNT
                        strTemp = strTemp & .EACH_PTN_INSPECTION_TIME(intPTN_Index) & ","
                    Next intPTN_Index
                    For intPTN_Index = 1 To .PATTERN_COUNT - 1
                        strTemp = strTemp & .PATTERN_NAME(intPTN_Index) & ","
                    Next intPTN_Index
                    strTemp = strTemp & .PATTERN_NAME(.PATTERN_COUNT)
                Else
                    strTemp = strTemp & ","
                End If
                Print #intFileNum, strTemp
            End If
        End With
    Next intIndex
    
    strTemp = "DEFECT_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_LCD_Data_CATST(ByVal pPath As String, ByVal pFileName As String, pLCD_DATA As DEFECT_FILE_LCD_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "LCD_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,LIGHT_ON_SOURCE_GRADE,LIGHT_ON_SOURCE_REASONCODE,TOTAL_LIGHT_ON_DEFECT_COUNT"
    Print #intFileNum, strTemp
    
    With pLCD_DATA
        strTemp = .PANELID & "," & .LIGHT_ON_SOURCE_GRADE & "," & .LIGHT_ON_SOURCE_REASON_CODE & "," & .TOTAL_LIGHT_ON_DEFECT_COUNT
        Print #intFileNum, strTemp
    End With
    
    strTemp = "LCD_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Panel_PDS_Summay_CATST(ByVal pPath As String, ByVal pFileName As String, pPANEL_PDS_SUMMARY As DEFECT_FILE_PANEL_PDS_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_PDS_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,PARAMETER_NAME,AVG,MIN,MAX,STD,COUNT"
    Print #intFileNum, strTemp
    
    With pPANEL_PDS_SUMMARY
        strTemp = .PANELID & "," & .PARAMETER_NAME & "," & .AVG & "," & .MIN & "," & .MAX & "," & .STD & "," & .Count
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_PDS_SUMMARY_END"
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Header_CALOI(ByVal pPath As String, ByVal pFileName As String, pHEADER_DATA As DEFECT_FILE_HEADER)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "HEADER_BEGIN"
    Print #intFileNum, strTemp
    
    With pHEADER_DATA
        strTemp = "FILE_VERSION:" & .JPS_VERSION
        Print #intFileNum, strTemp
        
        strTemp = "FILE_CREATED_TIME:" & .FILE_CREATE_TIME
        Print #intFileNum, strTemp
        
        strTemp = "EQUIP_TYPE:" & .EQUIP_TYPE
        Print #intFileNum, strTemp
        
        strTemp = "EQ_ID:" & .EQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "SUBEQ_ID:" & .SUBEQ_ID
        Print #intFileNum, strTemp
        
        strTemp = "CONTENT:PANEL_DATA/EQP_PANEL_DATA/PANEL_SUMMARY/DEFECT_DATA/LCD_DATA/PANEL_PDS_SUMMARY"
        Print #intFileNum, strTemp
    End With
    
    strTemp = "HEADER_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum
    
End Sub

Private Sub Write_Panel_Data_CALOI(ByVal pPath As String, ByVal pFileName As String, pPANEL_DATA As DEFECT_FILE_PANEL_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,GLASS_TYPE,PRODUCT_ID,PROCESS_ID,RECIPE_ID,SALEORDER,CF_GLASS_ID,ARRAY_LOT_ID,ARRAY_GLASS_ID,CF_GLASS_INFO,TFT_REPAIR_GRADE, GROUPID"
    Print #intFileNum, strTemp
        
    With pPANEL_DATA
        strTemp = .PANELID & "," & .GLASS_TYPE & "," & .PRODUCT_ID & "," & .PROCESS_ID & "," & .RECIPE_ID & "," & .SALEORDER & ","
        strTemp = strTemp & .CF_GLASS_ID & "," & .ARRAY_LOT_ID & "," & .ARRAY_GLASS_ID & "," & .CF_GLASS_OX_INFO & "," & .TFT_PANEL_JUDGE & "," & .GROUP_ID
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum
    
End Sub

Private Sub Write_EQP_Panel_Data_CALOI(ByVal pPath As String, ByVal pFileName As String, pEQP_PANEL_DATA As DEFECT_FILE_EQP_PANEL_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "EQP_PANEL_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "RECIPE_NO,START_TIME,END_TIME,OPERATOR_ID,TACT_TIME,MAIN_DEFECT_CODE,TOTAL_POINT_DEF_CNT,OPERATION_MODE,INSPECTION_TIME,TRANSFER_TIME"
    Print #intFileNum, strTemp
    
    With pEQP_PANEL_DATA
        strTemp = .RECIPE_NO & "," & .START_TIME & "," & .END_TIME & "," & .OPERATOR_ID & "," & .TACT_TIME & "," & .MAIN_DEFECT_CODE & ","
        strTemp = strTemp & .TOTAL_POINT_DEFECT_COUNT & "," & .OPERATION_MODE & "," & .INSPECTION_TIME & "," & .TRANSFER_TIME
        Print #intFileNum, strTemp
    End With
    
    strTemp = "EQP_PANEL_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Panel_Summary_CALOI(ByVal pPath As String, ByVal pFileName As String, pPANEL_SUMMARY As DEFECT_FILE_PANEL_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,JUDGE_RANK,MAIN_DEF_CODE,DATA_TOTAL_PIXEL,GATE_TOTAL_PIXEL,DRIVE_TYPE,LIGHT_ON_PRE_GRADE,LIGHT_ON_TARGET_REASON_TYPE_E"
    Print #intFileNum, strTemp
    
    With pPANEL_SUMMARY
        strTemp = .PANELID & "," & .JUDGE_RANK & "," & .MAIN_DEFECT_CODE & "," & .DATA_TOTAL_PIXEL & "," & .GATE_TOTAL_PIXEL & "," & .DRIVE_TYPE & ","
        strTemp = strTemp & .LIGHT_ON_PRE_GRADE & "," & .LIGHT_ON_TARGET_REASON_TYPE
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_SUMMARY_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Defect_Data_CALOI(ByVal pPath As String, ByVal pFileName As String, ByVal pArray_Count As Integer, pDEFECT_DATA() As DEFECT_FILE_DEFECT_DATA, _
                                    ByVal pPRODUCTID As String, ByVal pPanelID As String)

    Dim PFCD_ADDRESS_DATA               As PFCD_ADDRESS_STRUCTURE
    
    Dim intFileNum                      As Integer
    Dim intIndex                        As Integer
    Dim intPTN_Index                    As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "DEFECT_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,DEFECT_NO,DEFECT_CODE,COLOR,PANEL_GRADE,GRAY_LEVEL,DEFECT_NAME,DATA_X1,GATE_Y1,DATA_X2,GATE_Y2,DATA_X3,GATE_Y3,"
    strTemp = strTemp & "PANEL_COORDINATE_X1,PANEL_COORDINATE_Y1,PANEL_COORDINATE_X2,PANEL_COORDINDATE_Y2,PANEL_COORDINATE_X3,PANEL_COORDINATE_Y3,"
    strTemp = strTemp & "GLASS_COORDINATE_X1,GLASS_COORDINATE_Y1,GLASS_COORDINATE_X2,GLASS_COORDINDATE_Y2,GLASS_COORDINATE_X3,GLASS_COORDINATE_Y3,"
    strTemp = strTemp & "MARK_TYPE,EACH_PATTERN_INSPECTION_TIME,PATTERN_NAME"
    Print #intFileNum, strTemp
    
    Call Get_PFCD_ADDRESS_DATA(PFCD_ADDRESS_DATA, pPRODUCTID, CInt(Right(pPanelID, 2)))
    For intIndex = 1 To pArray_Count
        With pDEFECT_DATA(intIndex)
            If .DEFECT_CODE <> "" Then
                If Trim(.DATA_X1) <> "" Then
                    .PANEL_COORDINATE_X1 = CDbl(.DATA_X1) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y1) <> "" Then
                    .PANEL_COORDINATE_Y1 = CDbl(.GATE_Y1) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If Trim(.DATA_X2) <> "" Then
                    .PANEL_COORDINATE_X2 = CDbl(.DATA_X2) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y2) <> "" Then
                    .PANEL_COORDINATE_Y2 = CDbl(.GATE_Y2) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If Trim(.DATA_X3) <> "" Then
                    .PANEL_COORDINATE_X3 = CDbl(.DATA_X3) * (PFCD_ADDRESS_DATA.W + PFCD_ADDRESS_DATA.B1) + PFCD_ADDRESS_DATA.XC
                End If
                If Trim(.GATE_Y3) <> "" Then
                    .PANEL_COORDINATE_Y3 = CDbl(.GATE_Y3) * (PFCD_ADDRESS_DATA.L + PFCD_ADDRESS_DATA.B2) + PFCD_ADDRESS_DATA.YC
                End If
                If PFCD_ADDRESS_DATA.SOURCE_DIRECTION = "V" Then
                    Select Case PFCD_ADDRESS_DATA.ORIGIN_LOCATION
                    Case "LT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X3)
                        End If
                    Case "LB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                        End If
                    Case "RT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.DATA_X3)
                        End If
                    Case "RB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.GATE_Y3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.DATA_X3)
                        End If
                    End Select
                Else
                    Select Case PFCD_ADDRESS_DATA.ORIGIN_LOCATION
                    Case "LT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y3)
                        End If
                    Case "LB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO + CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y3)
                        End If
                    Case "RT":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO - CDbl(.GATE_Y3)
                        End If
                    Case "RB":
                        If Trim(.DATA_X1) <> "" Then
                            .GLASS_COORDINATE_X1 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X1)
                        End If
                        If Trim(.GATE_Y1) <> "" Then
                            .GLASS_COORDINATE_Y1 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y1)
                        End If
                        If Trim(.DATA_X2) <> "" Then
                            .GLASS_COORDINATE_X2 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X2)
                        End If
                        If Trim(.GATE_Y2) <> "" Then
                            .GLASS_COORDINATE_Y2 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y2)
                        End If
                        If Trim(.DATA_X3) <> "" Then
                            .GLASS_COORDINATE_X3 = PFCD_ADDRESS_DATA.XO - CDbl(.DATA_X3)
                        End If
                        If Trim(.GATE_Y3) <> "" Then
                            .GLASS_COORDINATE_Y3 = PFCD_ADDRESS_DATA.YO + CDbl(.GATE_Y3)
                        End If
                    End Select
                End If
                strTemp = .PANELID & "," & .DEFECT_NO & "," & .DEFECT_CODE & "," & .COLOR & "," & .PANEL_GRADE & "," & .GRAY_LEVEL & "," & .DEFECT_NAME & ","
                strTemp = strTemp & .DATA_X1 & "," & .GATE_Y1 & "," & .DATA_X2 & "," & .GATE_Y2 & "," & .DATA_X3 & "," & .GATE_Y3 & ","
                strTemp = strTemp & .PANEL_COORDINATE_X1 & "," & .PANEL_COORDINATE_Y1 & "," & .PANEL_COORDINATE_X2 & "," & .PANEL_COORDINATE_Y2 & "," & .PANEL_COORDINATE_X3 & "," & .PANEL_COORDINATE_Y3 & ","
                strTemp = strTemp & .GLASS_COORDINATE_X1 & "," & .GLASS_COORDINATE_Y1 & "," & .GLASS_COORDINATE_X2 & "," & .GLASS_COORDINATE_Y2 & "," & .GLASS_COORDINATE_X3 & "," & .GLASS_COORDINATE_Y3 & ","
                strTemp = strTemp & .MARK_TYPE & ","
                If .PATTERN_COUNT > 0 Then
                    For intPTN_Index = 1 To .PATTERN_COUNT
                        strTemp = strTemp & .EACH_PTN_INSPECTION_TIME(intPTN_Index) & ","
                    Next intPTN_Index
                    For intPTN_Index = 1 To .PATTERN_COUNT - 1
                        strTemp = strTemp & .PATTERN_NAME(intPTN_Index) & ","
                    Next intPTN_Index
                    strTemp = strTemp & .PATTERN_NAME(.PATTERN_COUNT)
                Else
                    strTemp = strTemp & ","
                End If
                Print #intFileNum, strTemp
            End If
        End With
    Next intIndex
    
    strTemp = "DEFECT_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_LCD_Data_CALOI(ByVal pPath As String, ByVal pFileName As String, pLCD_DATA As DEFECT_FILE_LCD_DATA)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "LCD_DATA_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,LIGHT_ON_SOURCE_GRADE,LIGHT_ON_SOURCE_REASONCODE,TOTAL_LIGHT_ON_DEFECT_COUNT"
    Print #intFileNum, strTemp
    
    With pLCD_DATA
        strTemp = .PANELID & "," & .LIGHT_ON_SOURCE_GRADE & "," & .LIGHT_ON_SOURCE_REASON_CODE & "," & .TOTAL_LIGHT_ON_DEFECT_COUNT
        Print #intFileNum, strTemp
    End With
    
    strTemp = "LCD_DATA_END" & vbCrLf
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

Private Sub Write_Panel_PDS_Summay_CALOI(ByVal pPath As String, ByVal pFileName As String, pPANEL_PDS_SUMMARY As DEFECT_FILE_PANEL_PDS_SUMMARY)

    Dim intFileNum                      As Integer
    
    Dim strTemp                         As String
    
    intFileNum = FreeFile
    
    Open pPath & pFileName For Append As intFileNum
    
    strTemp = "PANEL_PDS_SUMMARY_BEGIN"
    Print #intFileNum, strTemp
    
    strTemp = "PANEL_ID,PARAMETER_NAME,AVG,MIN,MAX,STD,COUNT"
    Print #intFileNum, strTemp
    
    With pPANEL_PDS_SUMMARY
        strTemp = .PANELID & "," & .PARAMETER_NAME & "," & .AVG & "," & .MIN & "," & .MAX & "," & .STD & "," & .Count
        Print #intFileNum, strTemp
    End With
    
    strTemp = "PANEL_PDS_SUMMARY_END"
    Print #intFileNum, strTemp
    
    Close intFileNum

End Sub

