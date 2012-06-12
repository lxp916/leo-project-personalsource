Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()

    Dim strDate                 As String
    Dim strTime                 As String
    Dim strPath                 As String
    
    If App.PrevInstance = False Then
        Load frmStart_Progress
        frmStart_Progress.Show
        Call Progress_Bar(10, "Sub folder initialize")
        If Init_Folder = False Then
            End
        Else
            If Make_DB = False Then
                Call MsgBox("Local Database create fail.", vbOKOnly, "DB create fail.")
            End If
            
            Call Make_MES_DB
            Call Make_AUTO_ALARM_DB
            
            Call Init_Class
            
            Call SaveLog("Main", "JPS program start.")
            Call Progress_Bar(10, "JPS Main window loading")
            Load frmMain
            frmMain.Show
            
            Load frmSystem_Log
            
            strDate = Format(DATE, "YYYY") & "/" & Format(DATE, "MM") & "/" & Format(DATE, "DD")
            strTime = Format(TIME, "HH") & ":" & Format(TIME, "MM") & ":" & Format(TIME, "SS")
            frmMain.flxRUN_Info.TextMatrix(5, 1) = strDate & " " & strTime
            
            Call Get_File_From_Host("Path.csv", "Table")
            Call Set_Path_Data
            Call Get_File_From_Host("User.csv", "User")
            Call Decode_User_Data
            Call Standard_Files_Download
'            Call Get_File_From_Host("Auto alarm.csv", "Table")
'            Call Decode_Auto_Alarm
'
            Unload frmStart_Progress
        End If
    Else
        Call MsgBox("JPS Program is already executed.", vbOKOnly, "Execution Fail.")
        End
    End If
    
End Sub

Private Sub Set_Path_Data()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    Dim strPath_Name                As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    Dim intLine_Index               As Integer
    
    strPath = App.PATH & "\Env\Standard_Info\"
    strFileName = "Path.csv"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        intLine_Index = 0
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intLine_Index = intLine_Index + 1
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                If UCase(Left(strTemp, intPos - 1)) <> "EQTYPE" Then
                    Call Insert_Path_Data(UCase(Left(strTemp, intPos - 1)), Mid(strTemp, intPos + 1))
                End If
            Else
                If UCase(strTemp) <> "FIELD NAME" Then
                    Select Case intLine_Index
                    Case 1:
                        strPath_Name = "EQTYPE"
                    Case 2:
                        strPath_Name = "PFCD.PID"
                    Case 3:
                        strPath_Name = "RANK"
                    Case 4:
                        strPath_Name = "USER"
                    Case 5:
                        strPath_Name = "PATTERN LIST"
                    Case 9:
                        strPath_Name = "VERSION"
                    Case 10:
                        strPath_Name = "TA_HISTORY"
                    End Select
                    Call Insert_Path_Data(strPath_Name, strTemp)
                End If
            End If
        Wend
        
        Close intFileNum
        
        ENV.Init_Class
    End If
    
End Sub

Private Sub Insert_Path_Data(ByVal pPath_Name As String, ByVal pPath As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM FS_PATH_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
        
            strQuery = "UPDATE FS_PATH_DATA SET "
            If (pPath_Name = "EQTYPE") Or (pPath_Name = "PFCD.PID") Or (pPath_Name = "RANK") Or (pPath_Name = "USER") Or _
               (pPath_Name = "PATTERN LIST") Or (pPath_Name = "VERSION") Or (pPath_Name = "TA_HISTORY") Then
                strQuery = strQuery & pPath_Name & " = '" & pPath & "'"
                
                dbMyDB.Execute (strQuery)
            End If
        Else
            lstRecord.Close
            
            strQuery = "INSERT INTO FS_PATH_DATA VALUES ("
            Select Case pPath_Name
            Case "EQTYPE":
                strQuery = strQuery & "'" & pPath & "', '', '', '', '', '', '')"
            Case "PFCD.PID":
                strQuery = strQuery & "'', '" & pPath & "', '', '', '', '', '')"
            Case "RANK":
                strQuery = strQuery & "'', '', '" & pPath & "', '', '', '', '')"
            Case "USER":
                strQuery = strQuery & "'', '', '', '" & pPath & "', '', '', '')"
            Case "PATTERN LIST":
                strQuery = strQuery & "'', '', '', '', '" & pPath & "', '', '')"
            Case "VERSION":
                strQuery = strQuery & "'', '', '', '', '', '" & pPath & "', '')"
            Case "TA_HISTORY":
                strQuery = strQuery & "'', '', '', '', '', '', '" & pPath & "')"
            End Select
            
            dbMyDB.Execute (strQuery)
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Private Sub Decode_User_Data()

    Dim typUser_Data                As USER_DATA
    
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    
    strPath = App.PATH & "\Env\Standard_Info\"
    strFileName = "User.csv"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                If UCase(Left(strTemp, intPos - 1)) <> "CARDNUM" Then
                    With typUser_Data
                        .USER_ID = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .USER_NAME = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .USER_PW1 = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .ID_CARD_CODE = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        intPos = InStr(strTemp, ",")
                        .USER_PW2 = Left(strTemp, intPos - 1)
                        .USER_LEVEL = Mid(strTemp, intPos + 1)
                    End With
                    Call Insert_User_Data(typUser_Data)
                End If
            End If
        Wend
            
        Close intFileNum
    End If
        
End Sub

Private Sub Insert_User_Data(pUser_Data As USER_DATA)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "SELECT * FROM USER_DATA WHERE "
        strQuery = strQuery & "USER_ID = '" & pUser_Data.USER_ID & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            
            With pUser_Data
                strQuery = "UPDATE USER_DATA SET "
                strQuery = strQuery & "USER_NAME = '" & .USER_NAME & "', "
                strQuery = strQuery & "ID_CARD_CODE = '" & .ID_CARD_CODE & "', "
                strQuery = strQuery & "USER_PW1 = '" & .USER_PW1 & "', "
                strQuery = strQuery & "USER_PW2 = '" & .USER_PW2 & "', "
                strQuery = strQuery & "USER_LEVEL = '" & .USER_LEVEL & "' WHERE "
                strQuery = strQuery & "USER_ID = '" & .USER_ID & "'"
                
                dbMyDB.Execute (strQuery)
            End With
        Else
            lstRecord.Close
            
            With pUser_Data
                strQuery = "INSERT INTO USER_DATA VALUES ("
                strQuery = strQuery & "'" & .USER_ID & "', "
                strQuery = strQuery & "'" & .USER_NAME & "', "
                strQuery = strQuery & "'" & .ID_CARD_CODE & "', "
                strQuery = strQuery & "'" & .USER_PW1 & "', "
                strQuery = strQuery & "'" & .USER_PW2 & "', "
                strQuery = strQuery & "'" & .USER_LEVEL & "')"
                
                dbMyDB.Execute (strQuery)
            End With
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Private Function Init_Folder() As Boolean

    Dim strFolder                   As String
    
On Error GoTo ErrorHandler
    
    Call Progress_Bar(10, "Environment folder initialize")
    strFolder = App.PATH & "\Env\"
    If Dir(strFolder, vbDirectory) = "" Then
        MkDir App.PATH & "\Env\"
    End If
    
    Call Progress_Bar(10, "Log folder initialize")
    strFolder = App.PATH & "\Log\"
    If Dir(strFolder, vbDirectory) = "" Then
        MkDir App.PATH & "\Log\"
    End If
    
    Call Progress_Bar(10, "Database folder initialize")
    strFolder = App.PATH & "\DB\"
    If Dir(strFolder, vbDirectory) = "" Then
        MkDir App.PATH & "\DB\"
    End If
    
    Call Progress_Bar(10, "GRADE_INFO folder initialize")
    strFolder = App.PATH & "\GRADE_INFO\"
    If Dir(strFolder, vbDirectory) = "" Then
        MkDir App.PATH & "\GRADE_INFO\"
    End If
    
    Call Progress_Bar(10, "STANDARD_INFO folder initialize")
    strFolder = App.PATH & "\Env\STANDARD_INFO\"
    If Dir(strFolder, vbDirectory) = "" Then
        MkDir App.PATH & "\STANDARD_INFO\"
    End If
    
    Init_Folder = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox(strFolder & " create fail. JPS program can't start up.", vbOKOnly, "Folder Initialize Error.")
    
    Init_Folder = False
    
End Function

Private Function Make_DB() As Boolean

    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    
    Dim dbMyDB                      As Database
    
    Dim strMsg                      As String
    
On Error GoTo ErrorHandler

    Call Progress_Bar(10, "Database initialize")
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then         'If Parameter.mdb is not exist
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        Call Progress_Bar(10, "Table initialize")
        If PATH_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "PATH_DATA Table create success.")
        End If
                
        If USER_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "USER_DATA Table create success.")
        End If
        
        If VERSION_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "VERSION_DATA Table create success.")
        End If
        
        If PFCD_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "PFCD_DATA Table create success.")
        End If
        
        If PFCD_ADDRESS_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "PFCD_ADDRESS_DATA Table create success.")
        End If
        
        If EQTYPE_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "EQTYPE_DATA Table create success.")
        End If
        
        If FS_PATH_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "FS_PATH_DATA Table create success.")
        End If
        
        If DEFECT_CODE_HIDE_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "DEFECT_CODE_HIDE_DATA Table create success.")
        End If
        
        If PATTERN_LIST(dbMyDB) = True Then
            Call Progress_Bar(5, "PATTERN_LIST Table create success.")
        End If
        
        If DEVICE_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "DEVICE_DATA Table create success.")
        End If
        
        If DEFECT_LIST(dbMyDB) = True Then
            Call Progress_Bar(5, "DEFECT_LIST Table create success.")
        End If
        
        If USEFUL_DEFECT(dbMyDB) = True Then
            Call Progress_Bar(5, "USEFUL_DEFECT Table create success.")
        End If
        
        If DEFECT_TYPE_PRIORITY(dbMyDB) = True Then
            Call Progress_Bar(5, "DEFECT_TYPE_PRIORITY Table create success.")
        End If
        
        If USER_LOGON_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "USER_LOGON_DATA Table create success.")
        End If
        
        dbMyDB.Close
    Else
        Call Progress_Bar(30, "Database file create skip.")
    End If
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        Call Progress_Bar(10, "Table initialize")
        If PANEL_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "PANEL_DATA Table create success.")
        End If
        
        dbMyDB.Close
    Else
        Call Progress_Bar(30, "Database file create skip.")
    End If
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "RANK_temp.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        Call Progress_Bar(10, "RANK Database initialize")
        If RANK_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "RANK_DATA Table create success.")
        End If
        
        If GRADE_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "GRADE_DATA Table create success.")
        End If
        
        dbMyDB.Close
    End If
    
    Call Make_Standard_Info_DB
    
    Make_DB = True
    
    Exit Function
    
ErrorHandler:

    Make_DB = False
    
End Function

Public Sub Make_Standard_Info_DB()

    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    
    Dim dbMyDB                      As Database
    
    Dim strMsg                      As String
    
On Error GoTo ErrorHandler
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO_Temp.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        Call Progress_Bar(10, "STANDARD_INFO Database initialize")
        If ITEM_CONTROL(dbMyDB) = True Then
            Call Progress_Bar(5, "ITEM_CONTROL Table create success.")
        End If
        
        If DEFECT_CODE_HIDE(dbMyDB) = True Then
            Call Progress_Bar(5, "DEFECT_CODE_HIDE Table create success.")
        End If
        
        If Abnormal_Panel(dbMyDB) = True Then
            Call Progress_Bar(5, "ABNORMAL_PANEL Table create success.")
        End If
        
        If Abnormal_MES_DATA(dbMyDB) = True Then
            Call Progress_Bar(5, "ABNORMAL_PANEL Table create success.")
        End If
        
        If Assign_Grade(dbMyDB) = True Then
            Call Progress_Bar(5, "ASSIGN_GRADE Table create success.")
        End If
        
        If PreJudgeChangeGrade1(dbMyDB) = True Then
            Call Progress_Bar(5, "PRE_JUDGE_CHANGE_GRADE1 Table create success.")
        End If
        
        If PreJudgeChangeGrade2(dbMyDB) = True Then
            Call Progress_Bar(5, "PRE_JUDGE_CHANGE_GRADE2 Table create success.")
        End If
        
        If PreJudgeChangeGrade3(dbMyDB) = True Then
            Call Progress_Bar(5, "PRE_JUDGE_CHANGE_GRADE3 Table create success.")
        End If
        
        If PostJudgeOtherRule1(dbMyDB) = True Then
            Call Progress_Bar(5, "POST_JUDGE_OTHER_RULE1 Table create success.")
        End If
        
        If PostJudgeOtherRule2(dbMyDB) = True Then
            Call Progress_Bar(5, "POST_JUDGE_OTHER_RULE2 Table create success.")
        End If
        
        If PostJudgeOtherRule3(dbMyDB) = True Then
            Call Progress_Bar(5, "POST_JUDGE_OTHER_RULE3 Table create success.")
        End If
        
        If PostJudgeGradeChange1(dbMyDB) = True Then
            Call Progress_Bar(5, "POST_JUDGE_GRADE_CHANGE1 Table create success.")
        End If
        
        If PostJudgeGradeChange2(dbMyDB) = True Then
            Call Progress_Bar(5, "POST_JUDGE_GRADE_CHANGE2 Table create success.")
        End If
        
        If CheckPanelIDChangeGrade(dbMyDB) = True Then
            Call Progress_Bar(5, "CHECK_PANELID_CHANGE_GRADE Table create success.")
        End If
        
        If ChangeGrade(dbMyDB) = True Then
            Call Progress_Bar(5, "CHANGE_GRADE Table create success.")
        End If
        
        If ChangeGradeByDefectCode(dbMyDB) = True Then
            Call Progress_Bar(5, "CHANGE_GRADE_DEFECT_CODE Table create success.")
        End If
        
        If RepairPointTimes(dbMyDB) = True Then
            Call Progress_Bar(5, "REPAIR_POINT_TIMES Table create success.")
        End If
        
        If FlagChangeGrade(dbMyDB) = True Then
            Call Progress_Bar(5, "FLAG_CHANGE_GRADE Table create success.")
        End If
        
        If SKChange(dbMyDB) = True Then
            Call Progress_Bar(5, "SK_CHANGE Table create success.")
        End If
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

End Sub

Private Function ITEM_CONTROL(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("ITEM_CONTROL")
    
    Call Create_Field(dbTable_Name, "ITEM_NAME", cDB_TEXT, True, 30, False)
    Call Create_Field(dbTable_Name, "USES", cDB_TEXT, True, 1, False)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    ITEM_CONTROL = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("ITEM_CONTROL table create fail.", vbOKOnly, "DB Table create fail.")
    
    ITEM_CONTROL = False
    
End Function

Private Function DEFECT_CODE_HIDE(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_CODE_HIDE")
    
    Call Create_Field(dbTable_Name, "PATTERN_LIST", cDB_TEXT, True, 10, False)
    For intIndex = 1 To 10
        Call Create_Field(dbTable_Name, "DEFECT_CODE" & intIndex, cDB_TEXT, False, 5, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_CODE_HIDE = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEFECT_CODE_HIDE table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEFECT_CODE_HIDE = False
    
End Function

Private Function Abnormal_Panel(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("ABNORMAL_PANEL")
    
    Call Create_Field(dbTable_Name, "ALARM_TEXT", cDB_TEXT, True, 200, False)
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "SHOP", cDB_TEXT, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    Abnormal_Panel = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("ABNORMAL_PANEL table create fail.", vbOKOnly, "DB Table create fail.")
    
    Abnormal_Panel = False
    
End Function

Private Function Abnormal_MES_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("ABNORMAL_MES_DATA")
    
    Call Create_Field(dbTable_Name, "ALARM_TEXT", cDB_TEXT, True, 200, False)
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DEST_FAB", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "RMANO", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "OQCNO", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_PANEL_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_REASON_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "CELL_LINE_RESCUE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "CELL_REPAIR_JUDGE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "TFT_REPAIR_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "CF_PANEL_OX_INFORMATION", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PANEL_OWNER_TYPE", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_CF", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_TFT", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_LCD", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "GROUPID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "REPAIR_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "POLARIZER_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "X_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "Y_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "LCD_Q_TAB_LOT_GROUPID", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "SK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "CF_R_DEFECT_CODE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ODK_AK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "BPAM_REWORK_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LCD_BRIGHT_DOT_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PS_HEIGHT_ERR_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_INSPECTION_NG_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "SEAL_INSPECTION_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_CHECKER_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_DOOR_OPEN_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "JOB_JUDGE", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "JOB_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "INSP_JUDGE_DATA_EQ1", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "INSP_JUDGE_DATA_EQ2", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "INSP_JUDGE_DATA_EQ3", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "INSP_JUDGE_DATA_EQ4", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "INSP_JUDGE_DATA_EQ5", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "SAMPLING_SLOT_FLAG", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "NEED_GRINDING_FLAG", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "SMALL_MULTI_PANEL_FLAG", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "CST_SETTING_CODE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_FLAG_CODE", cDB_TEXT, False, 4, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    Abnormal_MES_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("ABNORMAL_MES_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    Abnormal_MES_DATA = False
    
End Function

Private Function Assign_Grade(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("ASSIGN_GRADE")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DESTINATION_FAB", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PRIORITY", cDB_INTEGER, False, 1, True)
    
    
    pMyDB.TableDefs.Append dbTable_Name
    
    Assign_Grade = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("Assign_Grade table create fail.", vbOKOnly, "DB Table create fail.")
    
    Assign_Grade = False
    
End Function

Private Function PreJudgeChangeGrade1(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PRE_JUDGE_CHANGE_GRADE1")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DATA", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "GATE", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIMIT_UPPER", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "LIMIT_BOTTOM", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "LIMIT_LEFT", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "LIMIT_RIGHT", cDB_INTEGER, False, 1, True)
    For intIndex = 1 To 10
        Call Create_Field(dbTable_Name, "DEFECT_CODE" & intIndex, cDB_TEXT, False, 5, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PreJudgeChangeGrade1 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PRE_JUDGE_CHANGE_GRADE1 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PreJudgeChangeGrade1 = False
    
End Function

Private Function PreJudgeChangeGrade2(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PRE_JUDGE_CHANGE_GRADE2")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DATA", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "GATE", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIMIT_COUNT", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "TOTAL_DIVISION", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DEVIDE_DIVISION", cDB_INTEGER, False, 1, True)
    For intIndex = 1 To 3
        Call Create_Field(dbTable_Name, "DEFECT_CODE" & intIndex, cDB_TEXT, False, 5, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PreJudgeChangeGrade2 = True

    Exit Function
    
ErrorHandler:

    Call MsgBox("PRE_JUDGE_CHANGE_GRADE2 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PreJudgeChangeGrade2 = False
    
End Function

Private Function PreJudgeChangeGrade3(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PRE_JUDGE_CHANGE_GRADE3")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "POINT_DEFECT_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "DEFECT_TYPE", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PreJudgeChangeGrade3 = True

    Exit Function
    
ErrorHandler:

    Call MsgBox("PRE_JUDGE_CHANGE_GRADE3 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PreJudgeChangeGrade3 = False
    
End Function

Private Function PostJudgeOtherRule1(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("POST_JUDGE_OTHER_RULE1")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "SOURCE_DEFECT_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "CELL_REPAIR_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PostJudgeOtherRule1 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("POST_JUDGE_OTHER_RULE1 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PostJudgeOtherRule1 = False
    
End Function

Private Function PostJudgeOtherRule2(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("POST_JUDGE_OTHER_RULE2")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PostJudgeOtherRule2 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("POST_JUDGE_OTHER_RULE2 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PostJudgeOtherRule2 = False
    
End Function

Private Function PostJudgeOtherRule3(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("POST_JUDGE_OTHER_RULE3")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PRE_LOSS_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PostJudgeOtherRule3 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("POST_JUDGE_OTHER_RULE3 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PostJudgeOtherRule3 = False
    
End Function

Private Function PostJudgeGradeChange1(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("POST_JUDGE_GRADE_CHANGE1")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DATA_LINE", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "LEFT_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "RIGHT_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_TFT", cDB_TEXT, False, 4, True)
        
    pMyDB.TableDefs.Append dbTable_Name
    
    PostJudgeGradeChange1 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("POST_JUDGE_GRADE_CHANGE1 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PostJudgeGradeChange1 = False
    
End Function

Private Function PostJudgeGradeChange2(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("POST_JUDGE_GRADE_CHANGE2")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "TFT_REPAIR_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CELL_LINE_RESCUE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
        
    pMyDB.TableDefs.Append dbTable_Name
    
    PostJudgeGradeChange2 = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("POST_JUDGE_GRADE_CHANGE2 table create fail.", vbOKOnly, "DB Table create fail.")
    
    PostJudgeGradeChange2 = False
    
End Function

Private Function CheckPanelIDChangeGrade(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("CHECK_PANELID_CHANGE_GRADE")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    For intIndex = 1 To 10
        Call Create_Field(dbTable_Name, "NO_CHANGE_GRADE" & intIndex, cDB_TEXT, False, 2, True)
    Next intIndex
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "FILE_NAME", cDB_TEXT, False, 30, True)

    pMyDB.TableDefs.Append dbTable_Name
    
    CheckPanelIDChangeGrade = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("CHECK_PANELID_CHANGE_GRADE table create fail.", vbOKOnly, "DB Table create fail.")
    
    CheckPanelIDChangeGrade = False
    
End Function

Private Function ChangeGrade(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("CHANGE_GRADE")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "MACHINE_TYPE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    For intIndex = 1 To 6
        Call Create_Field(dbTable_Name, "MACHINE_ID" & intIndex, cDB_TEXT, False, 3, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    ChangeGrade = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("CHANAGE_GRADE table create fail.", vbOKOnly, "DB Table create fail.")
    
    ChangeGrade = True
    
End Function

Private Function ChangeGradeByDefectCode(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("CHANGE_GRADE_DEFECT_CODE")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DRIVE_TYPE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_REASON_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    ChangeGradeByDefectCode = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("CHANGE_GRADE_DEFECT_CODE table create fail.", vbOKOnly, "DB Table create fail.")
    
    ChangeGradeByDefectCode = False
    
End Function

Private Function RepairPointTimes(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("REPAIR_POINT_TIMES")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "REPAIR_REWORK_COUNT", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    For intIndex = 1 To 5
        Call Create_Field(dbTable_Name, "DEFECT_CODE" & intIndex, cDB_TEXT, False, 5, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    RepairPointTimes = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("REPAIR_POINT_TIMES table create fail.", vbOKOnly, "DB Table create fail.")
    
    RepairPointTimes = False
    
End Function

Private Function FlagChangeGrade(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("FLAG_CHANGE_GRADE")
    
    Call Create_Field(dbTable_Name, "PRE_GRADE", cDB_TEXT, True, 2, False)
    Call Create_Field(dbTable_Name, "NEW_GRADe", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, False, 16, True)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DEST_FAB", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "RMANO", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "OQCNO", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_PANEL_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_REASON_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "CELL_LINE_RESCUE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "CELL_REPAIR_JUDGE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "TFT_REPAIR_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "CF_PANEL_OX_INFORMATION", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PANEL_OWNER_TYPE", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_CF", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_TFT", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_LCD", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "GROUPID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "REPAIR_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "POLARIZER_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "X_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "Y_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "LCD_Q_TAB_LOT_GROUPID", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "SK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "CF_R_DEFECT_CODE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ODK_AK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "BPAM_REWORK_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LCD_BRIGHT_DOT_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PS_HIGHT_ERR_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_INSPECTION_NG_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "SEAL_INSPECTION_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_CHECKER_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_DOOR_OPEN_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "JOB_JUDGE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "JOB_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "BURR_CHECK_JUDGE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "BEVELING_JUDGE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CLEANER_M_PORT_JUDGE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "TEST_CV_JUDGE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "SAMPLING_SLOT_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PROCESS_INPUT_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "NEED_GRINDING_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "MISALIGNMENT_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "SMALL_MULTI_PANEL_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "NO_MATCH_GLASS_IN_BC_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_FLAG_CODE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PANEL_NG_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CUT_FLAG", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    FlagChangeGrade = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("FLAG_CHANGE_GRADE table create fail.", vbOKOnly, "DB Table create fail.")
    
    FlagChangeGrade = False
    
End Function

Private Function SKChange(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("SK_CHANGE")
    
    Call Create_Field(dbTable_Name, "MACHINE_NAME", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "SK_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "RANK", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "SAMPLING_VALUE", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "NEW_SK_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "NEW_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    SKChange = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("SK_CHANGE table create fail.", vbOKOnly, "DB Table create fail.")
    
    SKChange = False
    
End Function

Private Function RANK_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
   
     '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("RANK_DATA")
    
    Call Create_Field(dbTable_Name, "RANK_DIVISION", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_NAME", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "DEFECT_DIVISION", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "DEFECT_TYPE", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "JUDGE_OR_NOT", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "USE_XY", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "DETAIL_DIVISION", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ACCUMULATION", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "ADDRESS_COUNT", cDB_TEXT, False, 1, True)
    
     '============Leo 2012.05.22 Add Rank Level Start
    Call Create_Field(dbTable_Name, "ODF", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PRIORITY", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "POP_UP", cDB_TEXT, False, 1, True)
    For intRankLevel = 0 To UBound(RankLevel)
        Call Create_Field(dbTable_Name, "RANK_" & RankLevel(intRankLevel), cDB_TEXT, False, 50, True)
    Next intRankLevel
                            

'    Call Create_Field(dbTable_Name, "RANK_Y", cDB_TEXT, False, 50, True)
'    Call Create_Field(dbTable_Name, "RANK_L", cDB_TEXT, False, 50, True)
'    Call Create_Field(dbTable_Name, "RANK_K", cDB_TEXT, False, 50, True)
'    Call Create_Field(dbTable_Name, "RANK_C", cDB_TEXT, False, 50, True)
'    Call Create_Field(dbTable_Name, "RANK_S", cDB_TEXT, False, 50, True)
    '============Leo 2012.05.22 Add Rank Level End
   
    
    pMyDB.TableDefs.Append dbTable_Name
    
    RANK_DATA = True

    Exit Function
    
ErrorHandler:

    Call MsgBox("RANK_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    RANK_DATA = False
    
End Function

Private Function GRADE_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("GRADE_DATA")
    
    Call Create_Field(dbTable_Name, "RANK_DIVISION", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "RANK", cDB_TEXT, False, 5, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    GRADE_DATA = True

    Exit Function
    
ErrorHandler:

    Call MsgBox("GRADE_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    GRADE_DATA = False
    
End Function

Private Function DEFECT_TYPE_PRIORITY(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_TYPE_PRIORITY")
    
    Call Create_Field(dbTable_Name, "DEFECT_TYPE", cDB_TEXT, True, 1, False)
    Call Create_Field(dbTable_Name, "DEFECT_PRIORITY", cDB_INTEGER, True, 1, False)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_TYPE_PRIORITY = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEFECT_TYPE_PRIORITY table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEFECT_TYPE_PRIORITY = False
    
End Function

Private Function PANEL_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PANEL_DATA")
    
    Call Create_Field(dbTable_Name, "KEYID", cDB_TEXT, True, 35, False)     'KEYID = PANELID + _ + YYYYMMDD + HHMMSS
    Call Create_Field(dbTable_Name, "TIME", cDB_TEXT, False, 10, True)      '00:00:00
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, True, 20, False)
    Call Create_Field(dbTable_Name, "PANEL_RANK", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PANEL_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PANEL_LOSSCODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "LOSSCODE_NAME", cDB_TEXT, False, 25, True)
    Call Create_Field(dbTable_Name, "USER_NAME", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "PANEL_TYPE1", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "PANEL_TYPE2", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "PATH", cDB_TEXT, True, 200, False)
    Call Create_Field(dbTable_Name, "FILENAME", cDB_TEXT, True, 40, False)
    Call Create_Field(dbTable_Name, "RUN_DATE", cDB_LONG, False, 1, True)
    Call Create_Field(dbTable_Name, "RUN_TIME", cDB_LONG, False, 1, True)
    Call Create_Field(dbTable_Name, "TACT_TIME", cDB_LONG, False, 1, True)
    pMyDB.TableDefs.Append dbTable_Name
    
    PANEL_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PANEL_DATA table create fail.", vbOKOnly, "DB Table create fail.")

    PANEL_DATA = False
    
End Function

Private Function PATH_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PATH_DATA")
    
    Call Create_Field(dbTable_Name, "PATH_NAME", cDB_TEXT, True, 50, False)
    Call Create_Field(dbTable_Name, "PATH_DATA", cDB_TEXT, True, 100, False)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PATH_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PATH_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    PATH_DATA = False
    
End Function

Private Function USER_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("USER_DATA")
    
    Call Create_Field(dbTable_Name, "USER_ID", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "USER_NAME", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "ID_CARD_CODE", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "USER_PW1", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "USER_PW2", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "USER_LEVEL", cDB_TEXT, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    USER_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("USER_DATA table create fail.", vbOKCancel, "DB Table create fail.")
    
    USER_DATA = False
    
End Function

Private Function VERSION_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("VERSION_DATA")
    
    Call Create_Field(dbTable_Name, "MACHINE_ID", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "JPS_VERSION", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "EQ_VERSION", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "JPS_NAME", cDB_TEXT, False, 11, True)
    Call Create_Field(dbTable_Name, "INSTALL_DAY", cDB_TEXT, False, 14, True)
    Call Create_Field(dbTable_Name, "USER", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "JPS_SETUP_PATH", cDB_TEXT, False, 3, True)
    Call Create_Field(dbTable_Name, "JPS_LOG_PATH", cDB_TEXT, False, 3, True)
    Call Create_Field(dbTable_Name, "JPS_SERVER_PATH", cDB_TEXT, False, 3, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    VERSION_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("VERSION_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    VERSION_DATA = False
    
End Function

Private Function PFCD_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PFCD_DATA")
    
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 4, False)
    Call Create_Field(dbTable_Name, "X_PIXEL_LENGTH", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "Y_PIXEL_LENGTH", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "DATA", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "GATE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "CSTC", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "MAX_PANEL", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "PANEL_TYPE", cDB_TEXT, False, 4, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PFCD_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PFCD_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    PFCD_DATA = False
    
End Function

Private Function PFCD_ADDRESS_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PFCD_ADDRESS")
    
    Call Create_Field(dbTable_Name, "PRODUCT_ID", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "PANEL_NO", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "W", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "L", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "B1", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "B2", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "XC", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "YC", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "XO", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "YO", cDB_DOUBLE, False, 1, True)
    Call Create_Field(dbTable_Name, "ORIGIN_LOCATION", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "SOURCE_DIRECTION", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PFCD_ADDRESS_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PFCD_ADDRESS_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    PFCD_ADDRESS_DATA = False
    
End Function

Private Function EQTYPE_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("EQTYPE_DATA")
    
    Call Create_Field(dbTable_Name, "PC_NAME", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "PC_IP", cDB_TEXT, False, 15, True)
    Call Create_Field(dbTable_Name, "MACHINE_NAME", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "EQ_MODEL", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "FS_DRIVE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "FS_IP", cDB_TEXT, False, 16, True)
    Call Create_Field(dbTable_Name, "FS_USER_NAME", cDB_TEXT, False, 10, True)
    Call Create_Field(dbTable_Name, "FS_PASSWORD", cDB_TEXT, False, 6, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    EQTYPE_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("EQTYPE_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    EQTYPE_DATA = False
    
End Function

Private Function FS_PATH_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("FS_PATH_DATA")
    
    Call Create_Field(dbTable_Name, "EQTYPE", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "PFCD_PID", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "RANK", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "USER", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "PATTERN LIST", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "VERSION", cDB_TEXT, False, 100, True)
    Call Create_Field(dbTable_Name, "TA_HISTORY", cDB_TEXT, False, 100, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    FS_PATH_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("FS_PATH_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    FS_PATH_DATA = False
    
End Function

Private Function DEFECT_CODE_HIDE_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_CODE_HIDE_DATA")
    
    Call Create_Field(dbTable_Name, "PTN_LIST", cDB_TEXT, False, 10, True)
    For intIndex = 1 To 10
        Call Create_Field(dbTable_Name, "DEFECT_CODE" & intIndex, cDB_TEXT, False, 5, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_CODE_HIDE_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEFECT_CODE_HIDE_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEFECT_CODE_HIDE_DATA = False
    
End Function

Private Function PATTERN_LIST(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PATTERN_LIST")
    
    Call Create_Field(dbTable_Name, "FILENAME", cDB_TEXT, True, 15, False)
    Call Create_Field(dbTable_Name, "PATTERN_CODE", cDB_TEXT, True, 3, False)
    Call Create_Field(dbTable_Name, "PATTERN_NAME", cDB_TEXT, False, 20, True)
    Call Create_Field(dbTable_Name, "DELAY_TIME", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "LEVEL", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DH", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DL", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "VGH", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "VGL", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "RESCUE_HIGH", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "RESCUE_LOW", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "VCOM", cDB_INTEGER, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PATTERN_LIST = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("PATTERN_LIST table create fail.", vbOKOnly, "DB Table create fail.")
    
    PATTERN_LIST = False
    
End Function

Private Function USER_LOGON_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("USER_LOGON_DATA")
    
    Call Create_Field(dbTable_Name, "LOGON_DATE", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "LOGON_TIME", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "USER_ID", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "USER_NAME", cDB_TEXT, False, 10, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    USER_LOGON_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("USER_LOGON_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    USER_LOGON_DATA = False
    
End Function

Private Function DEVICE_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEVICE_DATA")
    
    Call Create_Field(dbTable_Name, "PORT_NO", cDB_INTEGER, True, 0, False)
    Call Create_Field(dbTable_Name, "DEVICE_NAME", cDB_TEXT, True, 20, False)
    Call Create_Field(dbTable_Name, "DEVICE_STATE", cDB_TEXT, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEVICE_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEVICE_DATA table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEVICE_DATA = False
    
End Function

Private Function DEFECT_LIST(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_LIST")
    
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DEFECT_NAME", cDB_TEXT, True, 30, False)
    Call Create_Field(dbTable_Name, "DEFECT_KIND", cDB_TEXT, True, 1, False)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_LIST = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEFECT_LIST table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEFECT_LIST = False
    
End Function

Private Function USEFUL_DEFECT(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("USEFUL_DEFECT")
    
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DEFECT_NAME", cDB_TEXT, True, 30, False)
    Call Create_Field(dbTable_Name, "DEFECT_KIND", cDB_TEXT, True, 1, False)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    USEFUL_DEFECT = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("USERFUL_DEFECT table create fail.", vbOKOnly, "DB Table create fail.")
    
    USEFUL_DEFECT = False
    
End Function

Private Function DEVICE_ONLINE_HISTORY(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEVICE_ONLINE_HISTORY")
    
    Call Create_Field(dbTable_Name, "MACHINE_NAME", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "ONLINE_DATE", cDB_LONG, True, 0, False)
    Call Create_Field(dbTable_Name, "ONLINE_TIME", cDB_LONG, True, 0, False)
    Call Create_Field(dbTable_Name, "PANEL_DRIVE_TYPE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "RUN_MODE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "USER_INFO", cDB_TEXT, False, 10, True)
    
    DEVICE_ONLINE_HISTORY = True
    
    Exit Function
    
ErrorHandler:

    Call MsgBox("DEVICE_ONLINE_HISTORY table create fail.", vbOKOnly, "DB Table create fail.")
    
    DEVICE_ONLINE_HISTORY = False
    
End Function

Private Sub Create_Field(pTable As TableDef, ByVal pField As String, ByVal pType As Integer, ByVal pRequired As Boolean, ByVal pSize As Integer, ByVal pAllowZeroLength As Boolean)

    Dim dbField_Name                As Field

On Error GoTo ErrorHandler

    Call Progress_Bar(10, pField & " table create.")
    Set dbField_Name = pTable.CreateField(pField)
    
    With dbField_Name
        Select Case pType
        Case cDB_TEXT:
            .Type = dbText
            .Required = pRequired
            .Size = pSize
            .AllowZeroLength = pAllowZeroLength
        Case cDB_BOOLEAN:
            .Type = dbBoolean
            .Required = pRequired
        Case cDB_INTEGER:
            .Type = dbInteger
            .Required = pRequired
        Case cDB_LONG:
            .Type = dbLong
            .Required = pRequired
        Case cDB_DOUBLE:
            .Type = dbDouble
            .Required = pRequired
        End Select
    End With
    pTable.Fields.Append dbField_Name
    
    Exit Sub
    
ErrorHandler:

    Call MsgBox(pField & " field create fail.", vbOKOnly, "DB Field create fail.")
    
End Sub

Private Sub Init_Class()

    Call Progress_Bar(5, "System parameter object initialize.")
    Call ENV.Init_Class
    
    Call Progress_Bar(5, "Command queue object initialize.")
    Call QUEUE.Init_Class
    
    Call Progress_Bar(5, "Rank_Obj object initialize.")
    Call RANK_OBJ.Init_Timer
    
End Sub

Private Sub Progress_Bar(ByVal pValue As Integer, ByVal pModule As String)

    Dim intIndex                    As Integer
    
    frmStart_Progress.lblModule_Name.Caption = pModule
'    For intIndex = 1 To pValue
        frmStart_Progress.shpProgress.Width = frmStart_Progress.shpProgress.Width + pValue
'    Next intIndex
    
End Sub

Public Sub Make_MES_DB()

    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    
    Dim dbMyDB                      As Database
    
    strDB_Path = App.PATH & "\DB\"
    If Dir(strDB_Path, vbDirectory) = "" Then
        MkDir strDB_Path
    End If
    strDB_FileName = "Panel_Data_Temp.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        If DEFECT_DATA(dbMyDB) = False Then
            Call MsgBox("DEFECT_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If PATTERN_INSPECTION(dbMyDB) = False Then
            Call MsgBox("PATTERN_INSPECTION Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If DEFECT_COUNT_DATA(dbMyDB) = False Then
            Call MsgBox("DEFECT_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If CST_MES_DATA(dbMyDB) = False Then
            Call MsgBox("CST_MES_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If PANEL_MES_DATA(dbMyDB) = False Then
            Call MsgBox("PANEL_MES_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If JOB_MES_DATA(dbMyDB) = False Then
            Call MsgBox("JOB_MES_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
        
        If SHARED_MES_DATA(dbMyDB) = False Then
            Call MsgBox("SHARE_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
    End If
    
End Sub

Private Function CST_MES_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("CST_MES_DATA")
    
    Call Create_Field(dbTable_Name, "CSTID", cDB_TEXT, True, 8, False)
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "OWNER", cDB_TEXT, True, 4, False)
    Call Create_Field(dbTable_Name, "PROCESSNUM", cDB_TEXT, True, 4, False)
    Call Create_Field(dbTable_Name, "PORTID", cDB_TEXT, True, 2, False)
    Call Create_Field(dbTable_Name, "PORTTYPE", cDB_TEXT, True, 2, False)
    Call Create_Field(dbTable_Name, "DEST_FAB", cDB_TEXT, True, 6, False)
    Call Create_Field(dbTable_Name, "PANELCOUNT", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "RMANO", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "OQCNO", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "SOURCE_FAB", cDB_TEXT, True, 6, False)
    For intIndex = 1 To 5
        Call Create_Field(dbTable_Name, "CST_SPARE" & intIndex, cDB_TEXT, False, 25, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    CST_MES_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("CST_MES_DATA", "CST_MES_DATA Table create fail.")
    
    CST_MES_DATA = False
    
End Function

Private Function PANEL_MES_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PANEL_MES_DATA")
    
    Call Create_Field(dbTable_Name, "SLOTNUM", cDB_TEXT, True, 4, False)
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, True, 12, False)
    Call Create_Field(dbTable_Name, "LIGHT_ON_PANEL_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_REASON_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "CELL_LINE_RESCUE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "CELL_REPAIR_JUDGE_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "TFT_REPAIR_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PANELID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "CF_PANEL_OX_INFORMATION", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PANEL_OWNER_TYPE", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_CF", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_TFT", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_LCD", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "GROUPID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "REPAIR_REWORK_COUNT", cDB_TEXT, False, cSIZE_REWORKCOUNT_MES, True)
    Call Create_Field(dbTable_Name, "CARBONIZATION_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "CARBONIZATION_GRADE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CARBONIZATION_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "POLARIZER_REWORK_COUNT", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "X_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "Y_TOTAL_PIXEL", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "X_ONE_PIXEL_LENGTH", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "Y_ONE_PIXEL_LENGTH", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "LCD_Q_TAB_LOT_GROUPID", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "SK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "CF_R_DEFECT_CODE", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "ODK_AK_FLAG", cDB_TEXT, False, 4, True)
    Call Create_Field(dbTable_Name, "BPAM_REWORK_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LCD_BRIGHT_DOT_FLAG", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "CF_PS_HEIGHT_ERR_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_INSPECTION_NG_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "PI_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_BAKE_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "HVA_OVER_Q_TIME_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "SEAL_INSPECTION_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_CHECKER_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "ODF_DOOR_OPEN_FLAG", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "LOT1_OPERATION_MODE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LOT2_OPERATION_MODE", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "PRODUCTID", cDB_TEXT, False, 12, True)
    Call Create_Field(dbTable_Name, "OWNERID", cDB_TEXT, False, 6, True)
    Call Create_Field(dbTable_Name, "PREPROCESSID", cDB_TEXT, False, 4, True)
    For intIndex = 1 To 10
        Call Create_Field(dbTable_Name, "PANEL_SPARE" & intIndex, cDB_TEXT, False, 25, True)
    Next intIndex
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PANEL_MES_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("PANEL_MES_DATA", "PANEL_MES_DATA Table create fail.")
    
    PANEL_MES_DATA = False
    
End Function

Private Function JOB_MES_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
    Dim intIndex                    As Integer
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("JOB_DATA")

    Call Create_Field(dbTable_Name, "CST_SEQUENCE", cDB_TEXT, False, cSIZE_CST_SEQUENCE_JOB, True)
    Call Create_Field(dbTable_Name, "JOB_SEQUENCE", cDB_TEXT, False, cSIZE_JOB_SEQUENCE_JOB, True)
    Call Create_Field(dbTable_Name, "CIM_MODE", cDB_TEXT, False, cSIZE_CIM_MODE_JOB, True)
    Call Create_Field(dbTable_Name, "JOB_JUDGE", cDB_TEXT, False, cSIZE_JOB_JUDGE_JOB, True)
    Call Create_Field(dbTable_Name, "JOB_GRADE", cDB_TEXT, False, cSIZE_JOB_GRADE_JOB, True)
    Call Create_Field(dbTable_Name, "GLASSID", cDB_TEXT, False, cSIZE_GLASSID_JOB, True)
    Call Create_Field(dbTable_Name, "BURR_CHECK_JUDGE", cDB_TEXT, False, cSIZE_BURR_CHECK_JUDGE_JOB, True)
    Call Create_Field(dbTable_Name, "BEVELING_JUDGE", cDB_TEXT, False, cSIZE_BEVELING_JUDGE_JOB, True)
    Call Create_Field(dbTable_Name, "CLEANER_M_PORT_JUDGE", cDB_TEXT, False, cSIZE_CLEANER_M_PORT_JUDGE_JOB, True)
    Call Create_Field(dbTable_Name, "TEST_CV_JUDGE", cDB_TEXT, False, cSIZE_TEST_CV_JUDGE_JOB, True)
    Call Create_Field(dbTable_Name, "SAMPLING_SLOT_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "PROCESS_INPUT_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "NEED_GRINDING_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "MISALIGNMENT_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "SMALL_MULTI_PANEL_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "AK_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "SK_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "NO_MATCH_GLASS_IN_BC_FLAG", cDB_TEXT, False, cSIZE_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "CASSETTE_SETTING_CODE", cDB_TEXT, False, cSIZE_CASSETTE_SETTING_CODE_JOB, True)
    Call Create_Field(dbTable_Name, "ABNORMAL_FLAG_CODE", cDB_TEXT, False, cSIZE_ABNORMAL_FLAG_CODE_JOB, True)
    Call Create_Field(dbTable_Name, "LIGHT_ON_REASON_CODE", cDB_TEXT, False, cSIZE_LIGHT_ON_REASON_CODE_JOB, True)
    Call Create_Field(dbTable_Name, "PANEL_NG_FLAG", cDB_TEXT, False, cSIZE_PANEL_NG_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "CUT_FLAG", cDB_TEXT, False, cSIZE_CUT_FLAG_JOB, True)
    Call Create_Field(dbTable_Name, "RESERVED", cDB_TEXT, False, cSIZE_RESERVED_JOB, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    JOB_MES_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("JOB_MES_DATA", "JOB_MES_DATA Table create fail.")
    
    JOB_MES_DATA = False
    
End Function

Private Function SHARED_MES_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("SHARED_DATA")
    
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, False, cSIZE_PANELID_SHARE, True)
    Call Create_Field(dbTable_Name, "GLASS_TYPE", cDB_TEXT, False, cSIZE_GLASS_TYPE_SHARE, True)
    Call Create_Field(dbTable_Name, "PRODUCTID", cDB_TEXT, False, cSIZE_PRODUCTID_SHARE, True)
    Call Create_Field(dbTable_Name, "PROCESSID", cDB_TEXT, False, cSIZE_PROCESSID_SHARE, True)
    Call Create_Field(dbTable_Name, "RECIPEID", cDB_TEXT, False, cSIZE_RECIPEID_SHARE, True)
    Call Create_Field(dbTable_Name, "SALE_ORDER", cDB_TEXT, False, cSIZE_SALE_ORDER_SHARE, True)
    Call Create_Field(dbTable_Name, "CF_GLASSID", cDB_TEXT, False, cSIZE_CF_GLASSID_SHARE, True)
    Call Create_Field(dbTable_Name, "ARRAY_LOTID", cDB_TEXT, False, cSIZE_ARRAY_LOTID_SHARE, True)
    Call Create_Field(dbTable_Name, "ARRAY_GLASSID", cDB_TEXT, False, cSIZE_ARRAY_GLASSID_SHARE, True)
    Call Create_Field(dbTable_Name, "CF_GLASS_INFO", cDB_TEXT, False, cSIZE_CF_GLASS_INFO_SHARE, True)
    Call Create_Field(dbTable_Name, "TFT_PANEL_JUDGE", cDB_TEXT, False, cSIZE_TFT_PANEL_JUDGE_SHARE, True)
    Call Create_Field(dbTable_Name, "PRE_PROCESSID1", cDB_TEXT, False, cSIZE_PRE_PROCESSID1_SHARE, True)
    Call Create_Field(dbTable_Name, "GROUPID", cDB_TEXT, False, cSIZE_GROUPID_SHARE, True)
    Call Create_Field(dbTable_Name, "TRANSFER_TIME", cDB_TEXT, False, cSIZE_TRANSFER_TIME_SHARE, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    SHARED_MES_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("SHARED_MES_DATA", "SHARED_DATA Table create fail.")
    
    SHARED_MES_DATA = False
End Function

Private Function DEFECT_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_DATA")
    
    Call Create_Field(dbTable_Name, "PANELID", cDB_TEXT, True, 20, False)
    Call Create_Field(dbTable_Name, "DEFECT_NO", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_NAME", cDB_TEXT, False, 8, True)
    Call Create_Field(dbTable_Name, "DETAIL_DIVISION", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "COLOR", cDB_TEXT, False, 1, True)
    Call Create_Field(dbTable_Name, "GRAY_LEVEL", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DEFECT_GATE1", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_DATA1", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_GATE2", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_DATA2", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_GATE3", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_DATA3", cDB_TEXT, False, 5, True)
    Call Create_Field(dbTable_Name, "DEFECT_TOTAL", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DEFECT_GRADE", cDB_TEXT, False, 2, True)
    
    pMyDB.TableDefs.Append dbTable_Name

    DEFECT_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("DEFECT_DATA", "DEFECT_DATA Table create fail.")
    
    DEFECT_DATA = False
    
    
End Function

Private Function PATTERN_INSPECTION(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("PATTERN_INSPECTION")
    
    Call Create_Field(dbTable_Name, "PATTERN_NAME", cDB_TEXT, False, 20, True)
    Call Create_Field(dbTable_Name, "PATTERN_START", cDB_LONG, False, 1, True)
    Call Create_Field(dbTable_Name, "PATTERN_END", cDB_LONG, False, 1, True)
    Call Create_Field(dbTable_Name, "INSPECTION_TIME", cDB_LONG, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    PATTERN_INSPECTION = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("PATTERN_INSPECTION", "PATTERN_INSPECTION Table create fail.")
    
    PATTERN_INSPECTION = False
    
End Function

Private Function DEFECT_COUNT_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("DEFECT_COUNT")
    
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DEFECT_COUNT", cDB_INTEGER, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_COUNT_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("DEFECT_COUNT_DATA", "DEFECT_COUNT_DATA Table create fail.")
    
    DEFECT_COUNT_DATA = False
    
End Function

Private Function DEFECT_LINE_LENGTH_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("LINE_LENGTH_DATA")
    
    Call Create_Field(dbTable_Name, "LINE_LENGTH_NO", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "DEFECT_DIVISION", cDB_TEXT, False, 2, True)
    Call Create_Field(dbTable_Name, "LINE_LENGTH", cDB_DOUBLE, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    DEFECT_LINE_LENGTH_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("DEFECT_LINE_LENGTH_DATA", "LINE_LENGTH_DATA Table create fail.")
    
    DEFECT_LINE_LENGTH_DATA = False
    
End Function

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Auto alarm database
'
'==========================================================================================================
Public Sub Make_AUTO_ALARM_DB()

    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    
    Dim dbMyDB                      As Database
    
    strDB_Path = App.PATH & "\DB\"
    If Dir(strDB_Path, vbDirectory) = "" Then
        MkDir strDB_Path
    End If
    strDB_FileName = "Auto_Alarm.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) = "" Then
        Set dbMyDB = Workspaces(0).CreateDatabase(strDB_Path & strDB_FileName, dbLangChineseSimplified)
        Set dbMyDB = OpenDatabase(strDB_Path & strDB_FileName)
        
        If AUTO_ALARM_DATA(dbMyDB) = False Then
            Call MsgBox("AUTO_ALARM_DATA Table create fail.", vbOKOnly, "DB Table create fail.")
        End If
    End If
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Auto alarm database table
'
'==========================================================================================================
Private Function AUTO_ALARM_DATA(pMyDB As Database) As Boolean

    Dim dbTable_Name                As TableDef
    
On Error GoTo ErrorHandler

    Set dbTable_Name = pMyDB.CreateTableDef("AUTO_ALARM_DATA")
    
    Call Create_Field(dbTable_Name, "PROCESS_NUM", cDB_TEXT, True, 4, False)
    Call Create_Field(dbTable_Name, "PFCD", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "DEFECT_CODE", cDB_TEXT, True, 5, False)
    Call Create_Field(dbTable_Name, "RANK", cDB_TEXT, True, 1, False)
    Call Create_Field(dbTable_Name, "COUNT_TIME", cDB_INTEGER, True, 1, False)
    Call Create_Field(dbTable_Name, "COUNT", cDB_INTEGER, True, 1, False)
    Call Create_Field(dbTable_Name, "ALARM_TEXT", cDB_TEXT, False, 200, True)
    Call Create_Field(dbTable_Name, "CURRENT_COUNT", cDB_INTEGER, False, 1, True)
    Call Create_Field(dbTable_Name, "EXPIRY_DATE", cDB_LONG, False, 1, True)
    Call Create_Field(dbTable_Name, "EXPIRY_TIME", cDB_LONG, False, 1, True)
    
    pMyDB.TableDefs.Append dbTable_Name
    
    AUTO_ALARM_DATA = True
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("AUTO_ALARM_DATA", "AUTO_ALARM_DATA Table create fail.")
    
    AUTO_ALARM_DATA = False
    
End Function

