Attribute VB_Name = "modCALOI_Sequence"
Option Explicit

Public Sub BLOI_Sequence(ByVal pPortID As Integer, ByVal pCommand As String)
    Dim typCST_INFO                             As CST_INFO_ELEMENTS
    Dim strDevice_State                         As String
    Dim strFileName                             As String
    Dim strLocalPath                            As String
    Dim LineX                                   As Integer
    Dim LineY                                   As Integer
    
    
    Dim intPortID                               As Integer
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortID, strDevice_State)
    
    If (strDevice_State = cDEVICE_ONLINE) Or (UCase(Left(pCommand, 4)) = "RONI") Then
        Select Case UCase(Left(pCommand, 4))
        Case "RONI":                                        'On line request from BLOI
            Call Decode_Online(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YONI")
            Call EQP.Set_LOI_STEP(cSTEP_RONI)
        Case "ROFI", "POFI":                                        'Off line request from BLOI
            Call State_Change(pPortID, Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), cDEVICE_OFFLINE)
            Call QUEUE.Put_Send_Command(pPortID, "YOFI")
        Case "RDCR":                                        'After Panel ID Read
            Call QUEUE.Put_Send_Command(pPortID, "YDCR")
'            Call Decode_PanelID_Read(pPortID, pCommand)
        Case "RBBC":                                        'Before block contact report from BLOI
            Call EQP.Set_RBBC_Command(pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YBBC")
        Case "RABC":                                        'After block contact report from BLOI
            Call EQP.Set_RABC_Command(pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YABC")
            Call EQP.Set_LOI_STEP(cSTEP_RABC)
        Case "RASO":
            Call QUEUE.Put_Send_Command(pPortID, "YASO")
        Case "PSON":                                        'Signal On from BLOI
            Call Receive_Signal_On_Reply(pPortID, pCommand)
        Case "PSOF":                                        'Signal Off from BLOI
            Call Recieve_Signal_Off_Reply(pPortID, pCommand)
        Case "RTRI":                                        'Manual trigger from BLOI
            Unload frmJudge
            Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QPPF")
            Call QUEUE.Put_Send_Command(pPortID, "YTRI")
            Select Case Decode_Manual_Trigger(pPortID, pCommand)
            Case 0:
                If (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Or (typCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Then
                 If Left(pubPANEL_INFO.OWNERID, 2) = "CD" Then
                       strFileName = UCase(Left(pubPANEL_INFO.OWNERID, 2) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
                 Else
                       strFileName = UCase(Left(pubCST_INFO.OWNER, 1) & Mid(pubCST_INFO.PFCD, 3, 5) & pubCST_INFO.PROCESS_NUM) & ".ran"
                 End If
                    strLocalPath = App.PATH & "\Env\Standard_Info\"
                    Call Get_File_From_Host(strFileName, strLocalPath)
                               'Lucas 2012.01.05 Ver.0.9.2 -----For CALOI use OWENERID=CD08 case
             '==========================================Start
                   
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
                
'Lucas Ver.1.9.34 2012.06.18=====For 1D1G address Show Alarm
    
If Mid(frmJudge.Text1, 2, 2) = "LD" Then
LineX = Val(frmJudge.Text2) / 2 + Val(frmJudge.Text4) / 2
LineY = Val(frmJudge.Text3) / 2 + Val(frmJudge.Text5) / 2
   If (Val(frmJudge.Text2) >= "1007" And Val(frmJudge.Text2) <= "1041") Or (Val(frmJudge.Text2) >= "2032" And Val(frmJudge.Text2) <= "2067") _
      Or (Val(frmJudge.Text2) >= "3056" And Val(frmJudge.Text2) <= "3090") Or (Val(frmJudge.Text2) >= "4080" And Val(frmJudge.Text2) <= "4113") _
      Or (LineX >= "1007" And LineX <= "1041") Or (LineX >= "2032" And LineX <= "2067") Or (LineX >= "3056" And LineX <= "3090") _
      Or (LineX >= "4080" And LineX <= "4113") Then
      If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
      Else
         If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
         End If
      End If
     
  End If
Else
    If Mid(frmJudge.Text1, 2, 2) = "LG" Then
     LineY = Val(frmJudge.Text3) / 2 + Val(frmJudge.Text5) / 2
     If (LineY >= "368" And LineY <= "400") Or (LineY >= "751" And LineY <= "783") _
         Or (Val(frmJudge.Text3) >= "368" And Val(frmJudge.Text3) <= "400") Or (Val(frmJudge.Text3) >= "751" And Val(frmJudge.Text3) <= "783") Then
       If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
        Else
          If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
          End If
       End If
    End If
   End If
End If

If Mid(frmJudge.Text6, 2, 2) = "LD" Then
LineX = Val(frmJudge.Text7) / 2 + Val(frmJudge.Text9) / 2

   If (Val(frmJudge.Text7) >= "1007" And Val(frmJudge.Text7) <= "1041") Or (Val(frmJudge.Text7) >= "2032" And Val(frmJudge.Text7) <= "2067") _
      Or (Val(frmJudge.Text7) >= "3056" And Val(frmJudge.Text7) <= "3090") Or (Val(frmJudge.Text7) >= "4080" And Val(frmJudge.Text7) <= "4113") _
      Or (LineX >= "1007" And LineX <= "1041") Or (LineX >= "2032" And LineX <= "2067") Or (LineX >= "3056" And LineX <= "3090") _
      Or (LineX >= "4080" And LineX <= "4113") Then
      If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
      Else
         If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
         End If
      End If
     
  End If
Else
    If Mid(frmJudge.Text6, 2, 2) = "LG" Then
     LineY = Val(frmJudge.Text8) / 2 + Val(frmJudge.Text10) / 2
     If (LineY >= "368" And LineY <= "400") Or (LineY >= "751" And LineY <= "783") _
        Or (Val(frmJudge.Text8) >= "368" And Val(frmJudge.Text8) <= "400") Or (Val(frmJudge.Text8) >= "751" And Val(frmJudge.Text8) <= "783") Then
       
       If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
        Else
          If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
          End If
       End If
    End If
   End If
End If

If Mid(frmJudge.Text11, 2, 2) = "LD" Then
LineX = Val(frmJudge.Text12) / 2 + Val(frmJudge.Text14) / 2

   If (Val(frmJudge.Text12) >= "1007" And Val(frmJudge.Text12) <= "1041") Or (Val(frmJudge.Text12) >= "2032" And Val(frmJudge.Text12) <= "2067") _
      Or (Val(frmJudge.Text12) >= "3056" And Val(frmJudge.Text12) <= "3090") Or (Val(frmJudge.Text12) >= "4080" And Val(frmJudge.Text12) <= "4113") _
      Or (LineX >= "1007" And LineX <= "1041") Or (LineX >= "2032" And LineX <= "2067") Or (LineX >= "3056" And LineX <= "3090") _
      Or (LineX >= "4080" And LineX <= "4113") Then
      If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
      Else
         If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Data线座标落在1D1G假线内,需判S/NG")
         End If
      End If
     
  End If
Else
    If Mid(frmJudge.Text11, 2, 2) = "LG" Then
     LineY = Val(frmJudge.Text13) / 2 + Val(frmJudge.Text15) / 2
     If (LineY >= "368" And LineY <= "400") Or (LineY >= "751" And LineY <= "783") _
        Or (Val(frmJudge.Text13) >= "368" And Val(frmJudge.Text13) <= "400") Or (Val(frmJudge.Text13) >= "751" And Val(frmJudge.Text13) <= "783") Then
       
       If pubCST_INFO.PROCESS_NUM = "3660" Then
'         strNew_Grade = "NG"
         Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
        Else
          If pubCST_INFO.PROCESS_NUM = "4660" Then
'           strNew_Grade = "S "
           Call Show_Message("1D1G", "该片LOI-1Gate线座标落在1D1G假线内,需判S/NG")
          End If
       End If
    End If
   End If
End If
'Lucas Ver.1.9.34 2012.06.18=====For 1D1G address Show Alarm
                
            Case 1:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0005CST_MES_DATA length error.")
            Case 2:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0006PANEL_MES_DATA length error.")
            Case 3:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0007JOB_MES_DATA length error.")
            Case 4:
                Call QUEUE.Put_Send_Command(pPortID, "QBAM0008SHARE_MES_DATA length error.")
            End Select
        Case "RBBU":                                        'Before block uncontact report from BLOI
            Call QUEUE.Put_Send_Command(pPortID, "YBBU")
        Case "REQS":                                        'EQ status report from BLOI
            Call Decode_EQ_Status_Report(pPortID, pCommand)
    '        Call QUEUE.Put_Send_Command(pPortID, "YEQS")
        Case "PEQS":                                        'EQ status reply from BLOI
            Call Decode_EQ_Status_Reply(pPortID, pCommand)
        Case "RRCO":                                        'Panel recontact report from BLOI
            Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QPPF")
            Call Receive_Panel_ReAlignment_Reply(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YRCO")
        Case "PBAM":                                        'Buzz and Message
            Call Receive_Buzz_Message_Reply(pPortID, pCommand)
        Case "PSPO":                                        'Panel out acknowledge from BLOI
            Call Recieve_Panel_Out_Reply(pPortID, pCommand)
            Call EQP.Set_LOI_STEP(cSTEP_PSPO)
        Case "PJPG":                                        'Judge panel grade acknowledge from BLOI
            Call Receive_Judge_Panel_Grade_Reply(pPortID, pCommand)
            Call EQP.Set_LOI_STEP(cSTEP_PJPG)
        Case "RJPG":                                        'Judge panel grade report from BLOI
            Call Decode_Judge_Panel_Grade(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YJPG")
        Case "RSCG":                                        'Setting change report form BLOI
            Call Decode_Setting_Change(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YSCG")
        Case "RVCR":                                        'After Panel ID Read from BLOI
            Call Decode_After_PanelID_Read(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YVCR")
        Case "RRAL":                                        'Panel Re-Alignment report from BLOI
            Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QPPF")
            Call Decode_Panel_ReAlignment(pPortID, pCommand)
            Call Receive_Panel_ReAlignment_Reply(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YRAL")
            Call EQP.Set_LOI_STEP(cSTEP_RRAL)
        Case "RSCO":                                        'Panel Shift Contact report from BLOI
            Call QUEUE.Put_Send_Command(EQP.Get_PG_PortID, "QPPF")
            Call Decode_Panel_Shift_Contact(pPortID, pCommand)
            Call Receive_Panel_ReAlignment_Reply(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YSCO")
        Case "RADC":                                        'Address Defect report from BLOI
            Call Decode_Address_Defect(pPortID, pCommand)
            Call QUEUE.Put_Send_Command(pPortID, "YADC")
        End Select
    End If
    
End Sub

Private Sub Decode_Online(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim LOGON_USER_DATA         As USER_LOGON_DATA
    Dim typVERSION_DATA         As VERSION_DATA
    
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
    frmMain.cmdForced_Unload.Visible = True
    frmMain.cmdForced_Unload.Enabled = True
    
    If Len(Mid(pCommand, 43)) > 10 Then
        strUSERINFO = Mid(pCommand, 43, 10)
    Else
        strUSERINFO = Mid(pCommand, 43)
    End If
    
    With frmMain.flxEQ_Information
        .TextMatrix(0, 1) = strTime
        .TextMatrix(1, 1) = strDriveType
        .TextMatrix(2, 1) = strPFCD
        .TextMatrix(3, 1) = strMode_State
        .TextMatrix(5, 1) = strMachineName
        Call ENV.Set_PG_Name("PG" & Right(.TextMatrix(5, 1), 3))
        Call ENV.Set_Current_Prober_Name(strMachineName)
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
    
'    frmMain.lblUser.Caption = strUSERINFO
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
'        strQuery = "INSERT INTO DEVICE_ONLINE_HISTORY VALUES ("
'        strQuery = strQuery & "'" & strMachineName & "', "
'        strQuery = strQuery & CLng(Left(strTIME, 8)) & ", "
'        strQuery = strQuery & CLng(Mid(strTIME, 9)) & ", "
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
'
    Call State_Change(pPortID, Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), cDEVICE_ONLINE)

'    Call ENV.Get_Device_Data_by_Name("API", intPortID, strDevice_State)
'    If strDevice_State <> cDEVICE_ONLINE Then
'        If intPortID <> 0 Then
'            strCommand = "QONA" & strTime & strDriveType & strPFCD & strRUNMode & strMachineName & strUSERINFO
'            intResult = QUEUE.Put_Send_Command(intPortID, strCommand)
'        End If
'    End If
    
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
    Dim strSubCommand           As String
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
        strSubCommand = Right(pCommand, 24)
        strPanelType1 = Left(strSubCommand, 12)
        strPanelType2 = Right(strSubCommand, 12)
                    
        Select Case strMESDataExistFlag
        Case "E":               'MES DATA enable & data exist
            strCST_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strCST_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSubCommand = Mid(pCommand, 1, CInt(strCST_Info_Length))
            If Len(strSubCommand) <> CInt(strCST_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                Decode_Manual_Trigger = 1
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSubCommand
            pCommand = Mid(pCommand, CInt(strCST_Info_Length) + 1)
            Call Decode_CST_Information_Elements(strSubCommand, typCST_INFO)
            
            If typCST_INFO.PFCD <> EQP.Get_Current_PFCD Then
                Call Get_File_From_Host("PFCD.PID", "Table")
                Call Read_PFCD_DATA
                Call Get_File_From_Host(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & Left(pubCST_INFO.OWNER, 1) & "address.csv", "Address")
                Call Read_PFCD_ADDRESS_DATA(Mid(pubCST_INFO.PFCD, 3, 5) & "_" & "address.csv")
            End If
            
            strPanel_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strMES_DATA_Command & strPanel_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSubCommand = Mid(pCommand, 1, CInt(strPanel_Info_Length))
            If Len(strSubCommand) <> CInt(strPanel_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 2
                End If
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSubCommand
            pCommand = Mid(pCommand, CInt(strPanel_Info_Length) + 1)
            Call Decode_PANEL_Information_Elements(strSubCommand, typPANEL_INFO, typCST_INFO.PFCD)
            
            strJob_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSubCommand = Mid(pCommand, 1, CInt(strJob_Info_Length))
            If Len(strSubCommand) <> CInt(strJob_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 3
                End If
            End If
            strJOB_DATA_Command = strSubCommand
            pCommand = Mid(pCommand, CInt(strJob_Info_Length) + 1)
            Call Decode_JOB_Information_Elements(strSubCommand, typJOB_INFO)
            
            strShare_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            'Lucas.Ver.0.9.13  2012.03.08=============================Add the Pcommand
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            '=========================================================End
            strSubCommand = Mid(pCommand, 1, CInt(strShare_Info_Length))
            If Len(strSubCommand) <> CInt(strShare_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
                If Decode_Manual_Trigger = 0 Then
                    Decode_Manual_Trigger = 4
                End If
            End If
            strSHARE_DATA_Command = strSubCommand
            pCommand = Mid(pCommand, CInt(strShare_Info_Length) + 1)
            Call Decode_Share_Information_Elements(strSubCommand, typSHARE_INFO)
            
            Call EQP.Set_MES_Data_for_API(strMESDataExistFlag, strJobDataExistFlag, strShareExistFlag, strMES_DATA_Command, strJOB_DATA_Command, strSHARE_DATA_Command)
            
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
                Call SaveLog("Decode_Manual_Trigger", typPANEL_INFO.PANELID & " data base create success.")
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
            
            If Mid(typCST_INFO.PFCD, 3, 5) <> Mid(EQP.Get_Current_PFCD, 3, 5) Or (typCST_INFO.PROCESS_NUM <> EQP.Get_Current_PROCESSID) Then
'                strDB_Path = App.PATH & "\DB\"
'                strDB_FileName = "STANDARD_INFO_Temp.mdb"
'                strDB_New_FileName = "STANDARD_INFO.mdb"
'
'                If Dir(strDB_Path & strDB_New_FileName, vbNormal) <> "" Then
'                    Kill strDB_Path & strDB_New_FileName
'                End If
'                FileCopy strDB_Path & strDB_FileName, strDB_Path & strDB_New_FileName
                
                strRemote_Path = ENV.Get_Path_Data("PATTERN LIST")
                strLocal_Path = App.PATH & "\Env\Standard_Info\"
                strFileName = UCase(Left(typCST_INFO.OWNER, 1) & Mid(typCST_INFO.PFCD, 3, 5) & typCST_INFO.PROCESS_NUM) & ".csv"
                Call Get_File_From_Host(strFileName, "Pattern")
                Call EQP.Read_PATTERN_LIST(strFileName)
                Call EQP.Set_PATTERN_LIST(strFileName)
                Call EQP.Set_Current_PFCD(pubCST_INFO.PFCD)
                Call EQP.Set_Current_PROCESSID(pubCST_INFO.PROCESS_NUM)
                
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
        
'        frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = ""
        frmMain.flxAlign_PanelID.TextMatrix(1, 0) = strPanelID
    End If
    
End Function

Private Sub Decode_PanelID_Read(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    Dim typJOB_INFO             As JOB_DATA_STRUCTURE
    Dim typSHARE_INFO           As SHARE_DATA_STRUCTURE
    Dim typPANEL_DATA           As PANEL_DATA
    
    Dim arrPanel_Type(1 To 2)   As String
    
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
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
    
    pCommand = Mid(pCommand, 5)
    strPanelID = Mid(pCommand, 1, cSIZE_PANELID)
    pCommand = Mid(pCommand, cSIZE_PANELID + 1)
    strMESDataExistFlag = Mid(pCommand, 1, (cSIZE_FLAG * 3))
    strJobDataExistFlag = Mid(strMESDataExistFlag, 2, 1)
    strShareExistFlag = Right(strMESDataExistFlag, 1)
    strMESDataExistFlag = Left(strMESDataExistFlag, 1)
    pCommand = Mid(pCommand, (cSIZE_FLAG * 3) + 1)
    strSUB_Command = Right(pCommand, 24)
    strPanelType1 = Left(strSUB_Command, 12)
    strPanelType2 = Right(strSUB_Command, 12)
    Select Case strMESDataExistFlag
    Case "E":               'MES DATA enable & data exist
            strCST_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strMES_DATA_Command = strCST_Info_Length
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strCST_Info_Length))
            If Len(strSUB_Command) <> CInt(strCST_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
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
            End If
            strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
            pCommand = Mid(pCommand, CInt(strPanel_Info_Length) + 1)
            Call Decode_PANEL_Information_Elements(strSUB_Command, typPANEL_INFO, typCST_INFO.PFCD)
            
            strJob_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            pCommand = Mid(pCommand, cSIZE_INFO_LENGTH + 1)
            strSUB_Command = Mid(pCommand, 1, CInt(strJob_Info_Length))
            If Len(strSUB_Command) <> CInt(strJob_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
            End If
            strJOB_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strJob_Info_Length) + 1)
            Call Decode_JOB_Information_Elements(strSUB_Command, typJOB_INFO)
            
            strShare_Info_Length = Mid(pCommand, 1, cSIZE_INFO_LENGTH)
            strSUB_Command = Mid(pCommand, 1, CInt(strShare_Info_Length))
            If Len(strSUB_Command) <> CInt(strShare_Info_Length) Then
                Call Show_Message("Data Error", "MES Data length error.")
            End If
            strSHARE_DATA_Command = strSUB_Command
            pCommand = Mid(pCommand, CInt(strShare_Info_Length) + 1)
            Call Decode_Share_Information_Elements(strSUB_Command, typSHARE_INFO)
            
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
            
'            strDB_Path = App.PATH & "\DB\"
'            strDB_FileName = "STANDARD_INFO_Temp.mdb"
'            strDB_New_FileName = "STANDARD_INFO.mdb"
'
'            If Dir(strDB_Path & strDB_New_FileName, vbNormal) <> "" Then
'                Kill strDB_Path & strDB_New_FileName
'            End If
'            FileCopy strDB_Path & strDB_FileName, strDB_Path & strDB_New_FileName
'
'            Call Standard_Files_Download
    Case "N":               'MES DATA enable & data not exist
    Case "D":               'MES DATA disable
    Case "S":               'In 1st inline light on
        strPFCD = Mid(pCommand, 1, cSIZE_PFCD)
        pCommand = Mid(pCommand, cSIZE_PFCD + 1)
        strOWNER = Mid(pCommand, cSIZE_OWNER)
    End Select

    frmMain.flxPre_Align_PanelID.TextMatrix(1, 0) = strPanelID

End Sub

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
    
    Call EQP.Set_DEFECT_UPLOAD(True)
    
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
    
    Call Standard_Files_Download

End Sub

Private Sub Decode_After_Signal_On(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strPanelID              As String
    Dim strMESDataExistFlag     As String
    Dim strWorkNo               As String
    Dim strPanelType1           As String
    Dim strPanelType2           As String
        
End Sub

Private Sub Receive_Signal_On_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strResponse             As String
    
'    strResponse = Mid(pCommand, 5, 1)
'
'    If strResponse = "1" Then
'    Else
'        Call Show_Message("Signal turn on fail.", "Signal on request fail received from EQ.")
'    End If
    
End Sub

Private Sub Recieve_Signal_Off_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim strResponse             As String
    
'    strResponse = Mid(pCommand, 5, 1)
'
'    If strResponse = "1" Then
'    Else
'        Call Show_Message("Signal turn off fail.", "Signal off request fail received from EQ.")
'    End If

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

Private Sub Decode_Panel_Recontact(ByVal pPortID As Integer, ByVal pCommand As String)

End Sub

Private Sub Receive_Panel_ReAlignment_Reply(ByVal pPortID As Integer, ByVal pCommand As String)

    Call EQP.set_Re_Alignment_Flag(True)
    
End Sub

Private Sub Decode_Panel_ReAlignment(ByVal pPortID As Integer, ByVal pCommand As String)

End Sub

Private Sub Decode_Panel_Shift_Contact(ByVal pPortID As Integer, ByVal pCommand As String)

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

Private Sub Decode_Address_Defect(ByVal pPortID As Integer, ByVal pCommand As String)

    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
    
    Dim strDATA_ADDRESS(1 To 3)         As String
    Dim strGATE_ADDRESS(1 To 3)         As String
        
    Dim strDefect_Code                  As String
    Dim strAddress_Count                As String
    
    Dim intDefect_Count                 As Integer
    Dim intIndex                        As Integer
    Dim intCol                          As Integer
    Dim intRow                          As Integer
    Dim intGrade_Count                  As Integer
    Dim intReceive_Address_Count        As Integer
        '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

   
    
    Call Reset_Interlock
    
    intRow = frmJudge.flxDefect_List.Rows - 1
    strDefect_Code = frmJudge.flxDefect_List.TextMatrix(intRow, 0)
    strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
    
    pCommand = Mid(pCommand, 5)
    
    'Address Count check
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
    
End Sub

