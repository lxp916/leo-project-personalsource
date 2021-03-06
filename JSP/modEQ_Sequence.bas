Attribute VB_Name = "modCATST_Sequence"
Option Explicit

Public Sub BTST_Sequence(ByVal pPortID As Integer, ByVal pCommand As String)
    Dim typCST_INFO                         As CST_INFO_ELEMENTS
    Dim strMode_State                       As String
    Dim strDevice_State                     As String
    Dim strMES_Exist                        As String
    Dim strJOB_Exist                        As String
    Dim strSHARE_Exist                      As String
    Dim strMES_DATA                         As String
    Dim strJOB_DATA                         As String
    Dim strSHARE_DATA                       As String
    Dim strCommand                          As String
    Dim strFileName                         As String
    Dim strLocalPath                        As String
    
    Dim intPortID                           As Integer
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortID, strDevice_State)
    
    If (strDevice_State = cDEVICE_ONLINE) Or (UCase(Left(pCommand, 4)) = "RONT") Then
        Select Case UCase(Left(pCommand, 4))
        Case "RONT":                                        'On line request from BTST
            Call Decode_Online(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YONT")
        Case "ROFT":                                        'Off line request from BTST
            Call State_Change(pPortID, Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), cDEVICE_OFFLINE)
            Call QUEUE.Put_Send_Command(pPortID, "YOFT")
        Case "POFT":                                        'Off line replyfrom BTST
            Call State_Change(pPortID, Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), cDEVICE_OFFLINE)
            Unload frmOffLine_Request
        Case "RBBC":                                        'Before block contact report from BTST
            Call EQP.Set_RBBC_Command(pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YBBC")
        Case "RABC":                                        'After block contact report from BTST
            Call EQP.Set_RABC_Command(pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YABC")
        Case "RASO":                                        'After Signal On from BTST
            Call Decode_After_Signal_On(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YASO")
        Case "RTRI":                                        'Manual trigger from BTST
            Unload frmJudge
            Select Case Decode_Manual_Trigger(pPortID, pCommand)
            Case 0:
                Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QPPF")
                Call QUEUE.Put_Send_Command(pPortID, "YTRI")
                
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
                If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
                    If strMode_State = "ON" Then
                        If (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Or (typCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Then
                            If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                                strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
                            Else
                                strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
                            End If
                            strLocalPath = App.PATH & "\Env\Standard_Info\"
                            Call Get_File_From_Host(strFileName, strLocalPath)
                            If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                                strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                            Else
                            strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".mdb"
                            End If
                            Call Read_Rank_Data(strFileName)
                        End If
                        Load frmJudge
                        Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QSMY" & Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5) & frmMain.flxMES_Data.TextMatrix(3, 1))
                        frmJudge.Show
                    End If
              End If
            Case 1:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0005CST_MES_DATA length error.")
            Case 2:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0006PANEL_MES_DATA length error.")
            Case 3:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0007JOB_MES_DATA length error.")
            Case 4:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0008SHARE_MES_DATA length error.")
            End Select
        Case "RBBU":                                        'Before block uncontact report from BTST
            Call QUEUE.Put_Send_Command(pPortID, "YBBU")
        Case "REQS":                                        'EQ status report from BTST
            Call Decode_EQ_Status_Report(pPortID, pCommand)
    '       Call QUEUE.Put_Send_Command(pPortID, "YEQS")
        Case "PEQS":                                        'EQ status reply from BTST
            Call Decode_EQ_Status_Reply(pPortID, pCommand)
        Case "PRCO":                                        'Panel recontact acknowledge from BTST
            Call Receive_Panel_Recontact_Reply(pPortID, pCommand)
        Case "PBAM":                                        'Buzz and Message
            Call Receive_Buzz_Message_Reply(pPortID, pCommand)
        Case "PSPO":                                        'Panel out acknowledge from BTST
            Call Recieve_Panel_Out_Reply(pPortID, pCommand)
        Case "PJPG":                                        'Judge panel grade acknowledge from BTST
            'Lucas 2012.02.15 Ver.0.9.7 ---------------For QSPO-->QJPG
             If frmMain.flxEQ_Information.TextMatrix(3, 1) <> "Full Auto" Then
             Call Receive_Judge_Panel_Grade_Reply(pPortID, pCommand)
             End If
        Case "RJPG":                                        'Judge panel grade report from BTST
            Call Decode_Judge_Panel_Grade(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YJPG")
        Case "RSCG":                                        'Setting change report form BTST
            Call Decode_Setting_Change(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YSCG")
        Case "RVCR":                                        'After Panel ID Read from BTST
            Call Decode_After_PanelID_Read(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YVCR")
        Case "PRAL":                                        'Panel Re-Alignment reply from BTST
            Call Receive_Panel_ReAlignment_Reply(pPortID, pCommand)
        Case "PSCO":                                        'Panel Shift Contact reply from BTST
            Call Receive_Panel_Shift_Contact_Reply(pPortID, pCommand)
        End Select
    End If
    
End Sub

Private Sub Decode_Online(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim LOGON_USER_DATA         As USER_LOGON_DATA
    Dim typVERSION_DATA         As VERSION_DATA
    Dim typPANEL_DATA           As DEFECT_FILE_PANEL_DATA
    
    Dim strTime                 As String
    Dim strDriveType            As String
    Dim strPFCD                 As String
    Dim strRUNMode              As String
    Dim strMode_State           As String
    Dim strMachineName          As String
    Dim strUSERINFO             As String
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strQuery                As String
    Dim strDevice_State         As String
    Dim strCommand              As String
    Dim strPath                 As String
    Dim strFileName             As String
    
    Dim intPortID               As Integer
    Dim intResult               As Integer
    
    Dim bolUser_Find            As Boolean
    
    Dim strMsg                  As String
    
On Error GoTo ErrorHandler

    strTime = Mid(pCommand, 5, cSIZE_TIME)
    strDriveType = Mid(pCommand, 19, cSIZE_PANELDRIVETYPE)
    strPFCD = Mid(pCommand, 21, cSIZE_PFCD)
    strRUNMode = Mid(pCommand, 33, cSIZE_RUNMODE)
    Select Case strRUNMode
    Case "ON":
        strMode_State = "Operator"
'        frmMain.cmdJudge.Enabled = True
    Case "IA":
        strMode_State = "Auto and RJS"
'        frmMain.cmdJudge.Enabled = False
    Case "FA":
        strMode_State = "Full Auto"
'        frmMain.cmdJudge.Enabled = False
    Case "EP":
        strMode_State = "EQ Pass"
'        frmMain.cmdJudge.Enabled = False
    End Select
    strMachineName = Mid(pCommand, 35, cSIZE_MACHINENAME)
    frmMain.cmdForced_Unload.Visible = False
    frmMain.cmdForced_Unload.Enabled = False
    
    If Len(Mid(pCommand, 43)) > 8 Then
        strUSERINFO = Mid(pCommand, 43, 8)
    Else
        strUSERINFO = Mid(pCommand, 43)
    End If
    
    With frmMain.flxEQ_Information
        .TextMatrix(0, 1) = strTime
        .TextMatrix(1, 1) = strDriveType
        .TextMatrix(2, 1) = strPFCD
        .TextMatrix(3, 1) = strMode_State
        .TextMatrix(5, 1) = strMachineName
        Call ENV.Set_Current_Prober_Name(strMachineName)
        Call ENV.Set_PG_Name("PG" & Right(.TextMatrix(5, 1), 3))
    End With
    
    With typVERSION_DATA
        .MACHINE_ID = ENV.Get_Current_Machine_Name
        .JPS_VERSION = App.Major & "." & App.Minor & "." & App.Revision
        .EQ_VERSION = ENV.Get_Current_Machine_Name
        .JPS_NAME = ENV.Get_JPS_Name
        strPath = App.PATH
        If Right(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
        strFileName = "JPS.exe"
        .INSTALL_DAY = Format(FileDateTime(strPath & strFileName), "YYYY-MM-DD")
        .JPS_SETUP_PATH = strPath
        .JPS_LOG_PATH = App.PATH & "\Log\"
        .JPS_SERVER_PATH = Get_Server_Path
    End With
    Call Set_Version_Data(typVERSION_DATA)
    
    frmMain.lblUser.Caption = strUSERINFO
    Call ENV.Set_Current_Machine_Name(strMachineName)
    
    If Trim(strUSERINFO) <> "" Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "Parameter.mdb"
        bolUser_Find = False
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "SELECT * FROM USER_DATA WHERE "
            strQuery = strQuery & "ID_CARD_CODE = '" & Trim(strUSERINFO) & "'"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                With LOGON_USER_DATA
                    .USER_ID = lstRecord.Fields("USER_ID")
                    .USER_NAME = lstRecord.Fields("USER_NAME")
                End With
                bolUser_Find = True
            Else
                bolUser_Find = False
            End If
            lstRecord.Close
            
            dbMyDB.Close
        Else
            bolUser_Find = False
        End If
    Else
        bolUser_Find = False
    End If
    
    Call EQP.Set_Re_Contact_Flag(False)
    Call EQP.set_Re_Alignment_Flag(False)
    
    If bolUser_Find = True Then
        Call ENV.Set_Current_Logon_User(LOGON_USER_DATA.USER_ID, LOGON_USER_DATA.USER_NAME)
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
        End Select
    Else
        Load frmLogin
        frmLogin.Show
    End If
    
'
'
'        strQuery = "INSERT INTO DEVICE_ONLINE_HISTORY VALUES ("
'        strQuery = strQuery & "'" & strMachineName & "', "
'        strQuery = strQuery & CLng(Left(strTime, 8)) & ", "
'        strQuery = strQuery & CLng(Mid(strTime, 9)) & ", "
'        strQuery = strQuery & "'" & strDriveType & "', "
'        strQuery = strQuery & "'" & strPFCD & "', "
'        strQuery = strQuery & "'" & strRUNMode & "', "
'        strQuery = strQuery & "'" & strUSERINFO & "')"
'
'        dbMyDB.Execute strQuery
'
'        dbMyDB.Close
'    Else
'        Call Show_Message("DB open fail", strDB_Path & strDB_FileName & " is not exist. Please close and restart JPS program.")
'    End If

    Call State_Change(pPortID, Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), cDEVICE_ONLINE)

    Call ENV.Get_Device_Data_by_Name("API", intPortID, strDevice_State)
    If strDevice_State <> cDEVICE_ONLINE Then
    'Lucas 2012.01.05 Ver.0.9.2----If PFCD is Space.Then it didn't Send Command
     If frmMain.flxEQ_Information.TextMatrix(2, 1) <> "            " Then
        If intPortID <> 0 Then
            strCommand = "QONA" & strTime & strDriveType & frmMain.flxEQ_Information.TextMatrix(2, 1) & strRUNMode & strMachineName & strUSERINFO
            intResult = QUEUE.Put_Send_Command(intPortID, strCommand)
        End If
     End If
    End If
    
    Exit Sub
    
ErrorHandler:

    strMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Decode_ONLINE", strMsg)
    
End Sub

Private Function Decode_Manual_Trigger(ByVal pPortID As Integer, ByVal pCommand As String) As Integer

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
    
    Decode_Manual_Trigger = 0
    
    If (EQP.Get_Re_Contact_Flag = False) And (EQP.Get_Re_Alignment_Flag = False) Then
        pCommand = Mid(pCommand, 5)
        strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
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
                Decode_Manual_Trigger = 1
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strCST_Info_Length) + 1)
            Call Decode_CST_Information_Elements(strSUB_Command, typCST_INFO)
            
            If typCST_INFO.PFCD <> EQP.Get_Current_PFCD Then
                Call Get_File_From_Host("PFCD.PID", "Table")
                Call Read_PFCD_DATA
                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & Left(pubCST_INFO.OWNER, 1) & "address.csv", "Address")
                Call Read_PFCD_ADDRESS_DATA(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv")
            End If
            
            strPanel_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strMES_DATA_Command & strPanel_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strPanel_Info_Length))
            If Len(strSUB_Command) <> CInt(strPanel_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 2
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
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 3
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
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 4
                End If
            End If
            strSHARE_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strShare_Info_Length) + 1)
            Call Decode_Share_Information_Elements(strSUB_Command, typSHARE_INFO)
            
            Call EQP.Set_MES_Data_for_API(strMESDataExistFlag, strJobDataExistFlag, strShareExistFlag, strMES_DATA_Command, strJOB_DATA_Command, strSHARE_DATA_Command)
            
            strRemote_Path = ENV.Get_Path_Data("PATTERN LIST")
            strLocal_Path = App.PATH & "\Env\Standard_Info\"
            strFileName = UCase(Left(typCST_INFO.OWNER, 1) & Mid(typCST_INFO.PFCD, 3, 5) & typCST_INFO.PROCESS_NUM) & ".csv"
            Call Get_File_From_Host(strFileName, "Pattern")
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
                If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CALOI" Then
                    If Len(pCommand) > 24 Then
                        .PANEL_TYPE1 = Left(pCommand, 12)
                        .PANEL_TYPE2 = Mid(pCommand, 13, 12)
                    End If
                End If
                .PATH = ""
                .FILENAME = ""
            End With
            If Make_CST_DATA_DB(Format(DATE, "MM"), Format(DATE, "DD"), typPANEL_DATA, typPANEL_INFO) = True Then
                Call SaveLog("Decode_Manual_Trigger", typPANEL_INFO.PANELID & " data base create success.")
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
'                strDB_Path = App.PATH & "\DB\"
'                strDB_FileName = "STANDARD_INFO_Temp.mdb"
'                strDB_New_FileName = "STANDARD_INFO.mdb"
'
'                If Dir(strDB_Path & strDB_New_FileName, vbNormal) <> "" Then
'                    Kill strDB_Path & strDB_New_FileName
'                End If
'                FileCopy strDB_Path & strDB_FileName, strDB_Path & strDB_New_FileName
                Call EQP.Set_Current_PFCD(pubCST_INFO.PFCD)
                Call EQP.Set_Current_PROCESSID(pubCST_INFO.PROCESS_NUM)
'                Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QSMY" & Mid(pubCST_INFO.PFCD, 3, 5))
'                Call Standard_Files_Download
            End If
            
            
    '        Call Decode_PANEL_Information_Elements(pCommand, typPANEL_INFO)
            Call Set_MES_Data(pubCST_INFO, typPANEL_INFO, typJOB_INFO, typSHARE_INFO)
        
                'TFT, CF Panel ID Check
            strMsg = Check_TFT_CF_PanelID(typPANEL_INFO.PANELID)
            If strMsg = "" Then
                'Check MES Data
                strMsg = Check_MES_Data(pubCST_INFO, pubPANEL_INFO, typJOB_INFO)
                If strMsg <> "" Then
                    Call Show_Message("Abnormal MES Data", strMsg)
                End If
            Else
                Call Show_Message("Abnormal Panel ID", strMsg)
            End If
        Case "N":               'MES DATA enable & data not exist
        Case "D":               'MES DATA disable
        Case "S":               'In 1st inline light on
            strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
            pCommand = Mid(pCommand, cSIZE_PFCD + 1)
            strOWNER = Mid(pCommand, cSIZE_OWNER)
        End Select
        
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = strPanelID
'        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = ""
    End If
    
End Function

Private Sub Decode_Judge_Panel_Grade(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strPanelID              As String
    Dim strGrade                As String
    Dim strLossCode             As String
    Dim strPanelJudgeMode       As String
    
    pCommand = Mid(pCommand, 5)
    strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
    pCommand = Mid(pCommand, cSIZE_PANELID + 1)
    
    strGrade = Mid(pCommand, 1, cSIZE_GRADE)
    pCommand = Mid(pCommand, cSIZE_GRADE + 1)
    
    strLossCode = Mid(pCommand, 1, cSIZE_LOSSCODE)
    pCommand = Mid(pCommand, cSIZE_LOSSCODE + 1)
    
    strPanelJudgeMode = Mid(pCommand, 1, cSIZE_PNL_JUDGEMODE)
    
    frmMain.lblPost_Judge.Caption = strGrade
    frmMain.lblPre_Loss_Code.Caption = strLossCode
    
    frmMain.flxRUN_Info.TextMatrix(1, 1) = CInt(frmMain.flxRUN_Info.TextMatrix(1, 1)) + 1
    
    Call EQP.Set_DEFECT_UPLOAD(False)
    
End Sub

Private Sub Decode_Setting_Change(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strTimeFlag             As String
    Dim strPanelDriveTypeFlag   As String
    Dim strPFCDFlag             As String
    Dim strRUNModeFlag          As String
    Dim strDate                 As String
    Dim strTime                 As String
    Dim strPanelDriveType       As String
    Dim strPFCD                 As String
    Dim strRUNMode              As String
    
    pCommand = Mid(pCommand, 5)
    
    strTimeFlag = Mid(pCommand, 1, cSIZE_FLAG)
    pCommand = Mid(pCommand, cSIZE_FLAG + 1)
    
    strPanelDriveTypeFlag = Mid(pCommand, 1, cSIZE_FLAG)
    pCommand = Mid(pCommand, cSIZE_FLAG + 1)
    
    strPFCDFlag = Mid(pCommand, 1, cSIZE_FLAG)
    pCommand = Mid(pCommand, cSIZE_FLAG + 1)
    
    strRUNModeFlag = Mid(pCommand, 1, cSIZE_FLAG)
    pCommand = Mid(pCommand, cSIZE_FLAG + 1)
    
    strTime = Mid(pCommand, 1, cSIZE_TIME)
    strDate = Left(strTime, 8)
    strTime = Mid(strTime, 9)
    pCommand = Mid(pCommand, cSIZE_TIME + 1)
    
    strPanelDriveType = Mid(pCommand, 1, cSIZE_PANELDRIVETYPE)
    pCommand = Mid(pCommand, cSIZE_PANELDRIVETYPE + 1)
    
    strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
    pCommand = Mid(pCommand, cSIZE_PFCD + 1)
    
    strRUNMode = Mid(pCommand, 1, cSIZE_RUNMODE)
    
End Sub

Private Sub Decode_EQ_Status_Report(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strRunningStatus        As String
    Dim strDCRStatus            As String
    
    pCommand = Mid(pCommand, 5)
    strRunningStatus = Mid(pCommand, 1, cSIZE_RUNNINGSTATUS)
    Select Case strRunningStatus
    Case "RNC":
        strRunningStatus = "Not contact"
    Case "RCO":
        strRunningStatus = "Contact"
    Case "REE":
        strRunningStatus = "Emergency error"
    Case "RSE":
        strRunningStatus = "Safty error"
    Case "RAE":
        strRunningStatus = "Auto alignment error"
    Case "RRE":
        strRunningStatus = "Auto Re-alignment error"
    Case "ROE":
        strRunningStatus = "Other error"
    End Select
    frmMain.flxEQ_Information.TextMatrix(4, 1) = strRunningStatus
    
'    pCommand = Mid(pCommand, cSIZE_RUNNINGSTATUS + 1)
'
'    strDCRStatus = Mid(pCommand, 1, cSIZE_DCRSTATUS)

End Sub

Private Sub Decode_EQ_Status_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strRunningStatus        As String
    Dim strDCRStatus            As String
    
    pCommand = Mid(pCommand, 5)
    strRunningStatus = Mid(pCommand, 1, cSIZE_RUNNINGSTATUS)
    Select Case strRunningStatus
    Case "RNC":
        strRunningStatus = "Not contact"
    Case "RCO":
        strRunningStatus = "Contact"
    Case "REE":
        strRunningStatus = "Emergency error"
    Case "RSE":
        strRunningStatus = "Safty error"
    Case "RAE":
        strRunningStatus = "Auto alignment error"
    Case "RRE":
        strRunningStatus = "Auto Re-alignment error"
    Case "ROE":
        strRunningStatus = "Other error"
    End Select
    frmMain.flxEQ_Information.TextMatrix(4, 1) = strRunningStatus
    
'    pCommand = Mid(pCommand, cSIZE_RUNNINGSTATUS + 1)
'
'    strDCRStatus = Mid(pCommand, 1, cSIZE_DCRSTATUS)
    
End Sub

Private Sub Decode_After_PanelID_Read(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    
    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
    Dim strPFCD                 As String
    Dim strOWNER                As String
    
    pCommand = Mid(pCommand, 5)
    strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
    pCommand = Mid(pCommand, cSIZE_PANELID + 1)
    strMESDataExistFlag = Mid(pCommand, 1, cSIZE_FLAG)
    pCommand = Mid(pCommand, cSIZE_FLAG + 1)
    Select Case strMESDataExistFlag
    Case "E":               'MES DATA enable & data exist
        strCST_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
        pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
        Call Decode_CST_Information_Elements(pCommand, typCST_INFO)
        strPanel_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
        pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
        Call Decode_PANEL_Information_Elements(pCommand, typPANEL_INFO, typCST_INFO.PFCD)
    Case "N":               'MES DATA enable & data not exist
    Case "D":               'MES DATA disable
    Case "S":               'In 1st inline light on
        strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
        pCommand = Mid(pCommand, cSIZE_PFCD + 1)
        strOWNER = Mid(pCommand, cSIZE_OWNER)
    End Select
    
End Sub

Private Sub Decode_After_Signal_On(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
        
End Sub

Private Sub Receive_Judge_Panel_Grade_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strCommand                  As String
    
    Dim intLength                   As Integer
    
    strCommand = "QSPO"
    
    intLength = cSIZE_PANELID - Len(Trim(frmMain.flxAlign_PanelID.TextMatrix(1, 0)))
    strCommand = strCommand & Trim(frmMain.flxAlign_PanelID.TextMatrix(1, 0)) & Space(intLength)
    
    Call QUEUE.Put_Send_Command(pPortID, strCommand)
    
End Sub

Private Sub Receive_Buzz_Message_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

End Sub

Private Sub Receive_Panel_Recontact_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim intPortNo                       As Integer
    
    Dim strStatus                       As String
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QDAA")
    End If
    Call EQP.Set_Re_Contact_Flag(True)
    
End Sub

Private Sub Receive_Panel_ReAlignment_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Call EQP.set_Re_Alignment_Flag(True)
    
End Sub

Private Sub Receive_Panel_Shift_Contact_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Call EQP.Set_Re_Contact_Flag(True)
    
End Sub

Private Sub Recieve_Panel_Out_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

'    Dim dbMyDB                  As Database
'
'    Dim strDB_Path              As String
'    Dim strDB_FileName          As String
'    Dim strQuery                As String
'    Dim strResponse             As String
'
'    Dim intRow                  As Integer
'
'    Call RANK_OBJ.Set_END_TIME(Format(TIME, "HHMMSS"))
'
'    strDB_Path = App.PATH & "\DB\"
'    strDB_FileName = "Result.mdb"
'
'    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
'        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
'
'        strQuery = "UPDATE PANEL_DATA SET "
'        strQuery = strQuery & "TACT_TIME=" & CLng(RANK_OBJ.Get_Tact_Time) & " WHERE "
'        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"
'
'        dbMyDB.Execute (strQuery)
'
'        dbMyDB.Close
'    End If
'
'    Call Make_Defect_File
    
    
'    strResponse = Mid(pCommand, 5, 1)
'    If strResponse = "1" Then
'        With frmMain
'            .flxAlign_PanelID.TextMatrix(1, 0) = ""
'            .cmbGrade_List.Text = "G"
'            .lblSelect_LossCode.Caption = ""
'            .lblPre_Judge.Caption = ""
'            .lblPost_Judge.Caption = ""
'            .lblPre_Loss_Code.Caption = ""
'            For intRow = 0 To .flxMES_Data.Rows - 1
'                .flxMES_Data.TextMatrix(intRow, 1) = ""
'            Next intRow
'        End With
'    Else
'        Call Show_Message("Panel out fail", "JPS received panel out not OK. Please check equipment.")
'    End If
    
End Sub

