VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim CURRENT_LOGON_USER          As USER_LOGON_DATA
Dim MESSAGE_DATA                As PUBLIC_NOTICE

Dim PATH_COUNT                  As Integer
Dim DEVICE_COUNT                As Integer
Dim MAX_RETRY_COUNT             As Integer
Dim ITEM_CONTROL_COUNT          As Integer

Dim PATH_NAME(1 To 7)           As String
Dim PATH_DATA(1 To 7)           As String

Dim PORT_USE(1 To 8)            As Boolean

Dim API_VERSION                 As String
Dim CURRENT_PROBER_NAME         As String
Dim CURRENT_MACHINE_NAME        As String
Dim JPS_NAME                    As String
Dim PG_NAME                     As String
Dim CURRENT_USER_LEVEL          As String
Dim CONTROL_UPDATE_TIME         As String
Dim DOWNLOAD_FLAG               As String

Dim DATA_CHANGE                 As Boolean

Dim DEVICE_DATA(1 To 8)         As DEVICE_INFO
Dim ITEM_CONTROL()              As ITEM_CONTROL

Public Sub Init_Class()

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strQuery                As String
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strPath_Name            As String
    Dim strPath_Data            As String
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    
    Dim intFileNum              As Integer
    Dim intRecord_Index         As Integer
    Dim intPortNo               As Integer
    
    PATH_COUNT = 0
    DEVICE_COUNT = 8
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM FS_PATH_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            PATH_COUNT = lstRecord.RecordCount
            lstRecord.MoveFirst
            intRecord_Index = 0
            For intRecord_Index = 1 To 7
                Select Case intRecord_Index
                Case 1:
                    strPath_Name = "EQTYPE"
                    strPath_Data = lstRecord.Fields("EQTYPE")
                Case 2:
                    strPath_Name = "PFCD.PID"
                    strPath_Data = lstRecord.Fields("PFCD_PID")
                Case 3:
                    strPath_Name = "RANK"
                    strPath_Data = lstRecord.Fields("RANK")
                Case 4:
                    strPath_Name = "USER"
                    strPath_Data = lstRecord.Fields("USER")
                Case 5:
                    strPath_Name = "PATTERN LIST"
                    strPath_Data = lstRecord.Fields("PATTERN LIST")
                Case 6:
                    strPath_Name = "VERSION"
                    strPath_Data = lstRecord.Fields("VERSION")
                Case 7:
                    strPath_Name = "TA_HISTORY"
                    strPath_Data = lstRecord.Fields("TA_HISTORY")
                End Select
                PATH_NAME(intRecord_Index) = strPath_Name
                PATH_DATA(intRecord_Index) = strPath_Data
            Next intRecord_Index
        End If
        lstRecord.Close
        PATH_COUNT = 7
        
'        strQuery = "SELECT * FROM DEVICE_DATA"
'
'        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
'
'        If lstRecord.EOF = False Then
'            lstRecord.MoveLast
'            DEVICE_COUNT = lstRecord.RecordCount
'            'ReDim DEVICE_DATA(DEVICE_COUNT)
'            lstRecord.MoveFirst
'            intRecord_Index = 0
'            While lstRecord.EOF = False
'                intPortNo = lstRecord.Fields("PORT_NO")
'                DEVICE_DATA(intPortNo).PORT_NO = lstRecord.Fields("PORT_NO")
'                DEVICE_DATA(intPortNo).DEVICE_NAME = lstRecord.Fields("DEVICE_NAME")
'                DEVICE_DATA(intPortNo).DEVICE_STATE = lstRecord.Fields("DEVICE_STATE")
'                If DEVICE_DATA(intPortNo).DEVICE_NAME = "BLOI" Then
'                    CURRENT_PROBER_NAME = "BLOI"
'                ElseIf DEVICE_DATA(intPortNo).DEVICE_NAME = "BTST" Then
'                    CURRENT_PROBER_NAME = "BTST"
'                End If
'                lstRecord.MoveNext
'            Wend
'        End If
'        lstRecord.Close
        dbMyDB.Close
    Else
        Call MsgBox("Parameter database file is not exist. Please restart program.", vbOKOnly, "Database file open error")
    End If
    CURRENT_MACHINE_NAME = ""
    With CURRENT_LOGON_USER
        .LOGON_TIME = ""
        .USER_ID = ""
        .USER_NAME = ""
    End With
    
    With MESSAGE_DATA
        .DATE = ""
        .UPDATED = False
        .MESSAGE_COUNT = 0
    End With
    
    strPath = App.PATH & "\Env\"
    strFileName = "DownloadFlag.dat"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        Line Input #intFileNum, strTemp
        
        Close intFileNum
        Call Set_Download_Flag(strTemp)
    Else
        strTemp = Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS")
        intFileNum = FreeFile
        
        Open strPath & strFileName For Output As intFileNum
        
        Print #intFileNum, strTemp
        
        Close intFileNum
        Call Set_Download_Flag(strTemp)
    End If
    
    Call Set_Port_Use
    
    DATA_CHANGE = True
    
End Sub

Public Sub Set_Port_Use()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    Dim intPort                     As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "PORT_USE.cfg"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                intPort = CInt(Left(strTemp, intPos - 1))
                If Mid(strTemp, intPos + 1) = "TRUE" Then
                    PORT_USE(intPort) = True
                Else
                    PORT_USE(intPort) = False
                End If
            End If
        Wend
        
        Close intFileNum
    End If
    
End Sub

Public Function Get_Port_Use(ByVal pPortID As Integer) As Boolean

    If (0 < pPortID) And (pPortID < 9) Then
        Get_Port_Use = PORT_USE(pPortID)
    Else
        Get_Port_Use = False
    End If
    
End Function

Public Function Get_Path_Count() As Integer

    Get_Path_Count = PATH_COUNT
    
End Function

Public Function Get_Path_Data(ByVal pPath_Name As String) As String

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    Get_Path_Data = ""
    bolFind = False
    intIndex = 0
    While (bolFind = False) And (intIndex < PATH_COUNT)
        intIndex = intIndex + 1
        If PATH_NAME(intIndex) = pPath_Name Then
            Get_Path_Data = PATH_DATA(intIndex)
            bolFind = True
        End If
    Wend
    
End Function

Public Sub Get_Path_Data_by_Index(ByVal pIndex As Integer, pPath_Name As String, pPath_Data As String)

    If (0 < pIndex) And (pIndex <= PATH_COUNT) Then
        pPath_Name = PATH_NAME(pIndex)
        pPath_Data = PATH_DATA(pIndex)
    Else
        pPath_Name = ""
        pPath_Data = ""
    End If
    
End Sub

Public Function Get_Device_Count() As Integer

    Get_Device_Count = DEVICE_COUNT
    
End Function

Public Sub Get_Device_Data_by_Name(ByVal pDevice_Name As String, pPort_No As Integer, pDevice_State As String)

    Dim intIndex                As Integer
    
    pPort_No = 0
    If DEVICE_COUNT > 0 Then
        For intIndex = 1 To DEVICE_COUNT
            If DEVICE_DATA(intIndex).DEVICE_NAME = pDevice_Name Then
                pPort_No = DEVICE_DATA(intIndex).PORT_NO
                pDevice_State = DEVICE_DATA(intIndex).DEVICE_STATE
            End If
        Next intIndex
    End If
    
End Sub

Public Sub Get_Device_Data_by_Index(ByVal pIndex As Integer, pPort_No As Integer, pDevice_Name As String, pDevice_State As String)

    If (0 < pIndex) And (pIndex < 9) Then
        If pIndex <= DEVICE_COUNT Then
            pPort_No = DEVICE_DATA(pIndex).PORT_NO
            pDevice_Name = DEVICE_DATA(pIndex).DEVICE_NAME
            pDevice_State = DEVICE_DATA(pIndex).DEVICE_STATE
        End If
    End If
    
End Sub

Public Sub Get_Device_Data_by_PortID(ByVal pPort_No As Integer, pDevice_Name As String, pDevice_State As String)

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    pDevice_Name = ""
    pDevice_State = ""
    If (0 < pPort_No) And (pPort_No < 9) Then
        pDevice_Name = DEVICE_DATA(pPort_No).DEVICE_NAME
        pDevice_State = DEVICE_DATA(pPort_No).DEVICE_STATE
    End If
    
End Sub

Public Sub Set_API_Version(ByVal pVersion As String)
    
    API_VERSION = pVersion
    
End Sub

Public Function Get_API_Version() As String

    Get_API_Version = API_VERSION
    
End Function

Public Sub Set_Max_Retry_Count(ByVal pCount As Integer)

    MAX_RETRY_COUNT = pCount
    
End Sub

Public Function Get_Max_Retry_Count() As Integer

    Get_Max_Retry_Count = MAX_RETRY_COUNT
    
End Function

Public Sub Set_Current_Prober_Name(ByVal pProber_Name As String)

    CURRENT_PROBER_NAME = pProber_Name
    JPS_NAME = "CAJPS" & Right(CURRENT_PROBER_NAME, 3)
    
End Sub

Public Function Get_Current_Prober_Name() As String

    Get_Current_Prober_Name = CURRENT_PROBER_NAME
    
End Function

Public Sub Set_Device_Info(ByVal pPortNo As Integer, ByVal pDevice_Name As String, ByVal pDevice_State As String)

    Dim dbMyDB                  As Database
    
    Dim lstRecord               As Recordset
    
    Dim strQuery                As String
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    
    Dim intRecord_Index         As Integer
    Dim intIndex                As Integer
    
    PATH_COUNT = 0
    DEVICE_COUNT = 8
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM DEVICE_DATA WHERE "
        strQuery = strQuery & "DEVICE_NAME = '" & pDevice_Name & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            
            strQuery = "UPDATE DEVICE_DATA SET "
            strQuery = strQuery & "PORT_NO = " & pPortNo & ", "
            strQuery = strQuery & "DEVICE_STATE = '" & pDevice_State & "' WHERE "
            strQuery = strQuery & "DEVICE_NAME = '" & pDevice_Name & "'"
            
            dbMyDB.Execute strQuery
        Else
            lstRecord.Close
            
            strQuery = "INSERT INTO DEVICE_DATA VALUES ("
            strQuery = strQuery & pPortNo & ", "
            strQuery = strQuery & "'" & pDevice_Name & "', "
            strQuery = strQuery & "'" & pDevice_State & "')"
            
            dbMyDB.Execute strQuery
        End If
        
        dbMyDB.Close
        
        For intIndex = 1 To 8
            If DEVICE_DATA(intIndex).DEVICE_NAME = pDevice_Name Then
                DEVICE_DATA(intIndex).DEVICE_NAME = ""
                DEVICE_DATA(intIndex).DEVICE_STATE = ""
                DEVICE_DATA(intIndex).PORT_NO = 0
            End If
        Next intIndex
        
        DEVICE_DATA(pPortNo).DEVICE_NAME = pDevice_Name
        DEVICE_DATA(pPortNo).DEVICE_STATE = pDevice_State
        DEVICE_DATA(pPortNo).PORT_NO = pPortNo
    End If
    
End Sub

Public Sub Set_Device_Info_by_Index(ByVal pPortNo As Integer, ByVal pDevice_Name As String, ByVal pDevice_State As String)

    If (0 < pPortNo) And (pPortNo < 9) Then
'        DEVICE_DATA(pPortNo).DEVICE_NAME = pDevice_Name
        DEVICE_DATA(pPortNo).DEVICE_STATE = pDevice_State
        DEVICE_DATA(pPortNo).PORT_NO = pPortNo
    End If
    
End Sub

Public Sub Set_Item_Control(pITEM_NAME() As String, pUSE() As String, ByVal pData_Count As Integer)

    Dim intIndex                        As Integer
    
    If pData_Count > 0 Then
        ReDim ITEM_CONTROL(pData_Count)
        For intIndex = 1 To pData_Count
            ITEM_CONTROL(intIndex).ITEM_NAME = pITEM_NAME(intIndex)
            ITEM_CONTROL(intIndex).ENABLE_DISABLE = pUSE(intIndex)
        Next intIndex
        ITEM_CONTROL_COUNT = pData_Count
    End If
    
End Sub

Public Function Get_Item_Control_Count() As Integer

    Get_Item_Control_Count = ITEM_CONTROL_COUNT
    
End Function

Public Sub Get_Item_Control_by_Index(ByVal pIndex As Integer, pITEM_NAME As String, pUSE As String)

    If (0 < pIndex) And (pIndex <= ITEM_CONTROL_COUNT) Then
        pITEM_NAME = ITEM_CONTROL(pIndex).ITEM_NAME
        pUSE = ITEM_CONTROL(pIndex).ENABLE_DISABLE
    End If
    
End Sub

Public Function Get_Item_Control_by_Name(ByVal pITEM_NAME As String) As String

    Dim intIndex                        As Integer
    
    Get_Item_Control_by_Name = ""
    For intIndex = 1 To ITEM_CONTROL_COUNT
        If ITEM_CONTROL(intIndex).ITEM_NAME = pITEM_NAME Then
            Get_Item_Control_by_Name = ITEM_CONTROL(intIndex).ENABLE_DISABLE
        End If
    Next intIndex
    
End Function

Public Sub Get_Item_Control(pITEM_NAME() As String, pUSE() As String)

    Dim intIndex                        As Integer
    
    For intIndex = 1 To ITEM_CONTROL_COUNT
        pITEM_NAME(intIndex) = ITEM_CONTROL(intIndex).ITEM_NAME
        pUSE(intIndex) = ITEM_CONTROL(intIndex).ENABLE_DISABLE
    Next intIndex
    
End Sub

Public Sub Set_Current_Machine_Name(ByVal pMachineName As String)

    CURRENT_MACHINE_NAME = pMachineName
    
End Sub

Public Function Get_Current_Machine_Name() As String

    Get_Current_Machine_Name = CURRENT_MACHINE_NAME
    
End Function

Public Sub Set_Current_Logon_User(ByVal pUSER_ID As String, ByVal pUSER_NAME As String)

    Dim dbMyDB                      As Database
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    If pUSER_ID <> "Lucas" Then
        With CURRENT_LOGON_USER
            .LOGON_DATE = Format(DATE, "YYYYMMDD")
            .LOGON_TIME = Format(TIME, "HHMMSS")
            .USER_ID = pUSER_ID
            .USER_NAME = pUSER_NAME
        
            strDB_Path = App.PATH & "\DB\"
            strDB_FileName = "Parameter.mdb"
            
            If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
                Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
                
                strQuery = "INSERT INTO USER_LOGON_DATA VALUES ("
                strQuery = strQuery & "'" & .LOGON_DATE & "', "
                strQuery = strQuery & "'" & .LOGON_TIME & "', "
                strQuery = strQuery & "'" & .USER_ID & "', "
                strQuery = strQuery & "'" & .USER_NAME & "')"
                
                dbMyDB.Execute strQuery
                
                dbMyDB.Close
            End If
        End With
        Call Set_User_Level
    Else
        CURRENT_USER_LEVEL = "S"
    End If
    
End Sub

Public Sub Get_Current_Logon_User(pLOGON_TIME As String, pUSER_ID As String, pUSER_NAME As String)

    With CURRENT_LOGON_USER
        pLOGON_TIME = .LOGON_TIME
        pUSER_ID = .USER_ID
        pUSER_NAME = .USER_NAME
    End With
    
End Sub

Public Sub Set_Notice_Count(ByVal pCount As Integer)

    If pCount > 0 Then
        ReDim MESSAGE_DATA.MESSAGE(pCount)
    End If
    MESSAGE_DATA.MESSAGE_COUNT = pCount
    
End Sub

Public Sub Set_Notice_Data(ByVal pIndex As Integer, ByVal pDATE As String, ByVal pMESSAGE As String, ByVal pTime As String)

    Dim intIndex                As Integer
    
    With MESSAGE_DATA
        If .MESSAGE_COUNT > 0 Then
            .DATE = pDATE
            .MESSAGE(pIndex) = pMESSAGE
            .TIME = pTime
            .UPDATED = True
        End If
    End With
    
End Sub

Public Sub Get_Notice_Data(pDATE As String, pMESSAGE() As String, pTime As String)

    Dim intIndex                As Integer
    
    With MESSAGE_DATA
        If .MESSAGE_COUNT > 0 Then
            pDATE = .DATE
            For intIndex = 1 To .MESSAGE_COUNT
                pMESSAGE(intIndex) = .MESSAGE(intIndex)
            Next intIndex
            pTime = .TIME
        End If
    End With
    
End Sub

Public Function Get_NOTICE_MESSAGE_by_Index(ByVal pIndex) As String

    If (0 < pIndex) And (pIndex <= MESSAGE_DATA.MESSAGE_COUNT) Then
        Get_NOTICE_MESSAGE_by_Index = MESSAGE_DATA.MESSAGE(pIndex)
    End If
    
End Function

Public Sub Set_NOTICE_UPDATE(ByVal pUPDATE As Boolean)

    MESSAGE_DATA.UPDATED = pUPDATE
    
End Sub

Public Function Get_NOTICE_UPDATE() As Boolean

    Get_NOTICE_UPDATE = MESSAGE_DATA.UPDATED
    
End Function

Public Function Get_NOTICE_Count() As Integer

    Get_NOTICE_Count = MESSAGE_DATA.MESSAGE_COUNT
    
End Function

Public Function Get_JPS_Name() As String

    Get_JPS_Name = JPS_NAME
    
End Function

Public Sub Set_PG_Name(ByVal pName As String)

    PG_NAME = pName
    
End Sub

Public Function Get_PG_Name() As String

    Get_PG_Name = PG_NAME
    
End Function

Public Sub Set_JPS_Name(ByVal pName As String)

    JPS_NAME = pName
    
End Sub

Public Sub Set_User_Level()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strCurrent_User             As String
    Dim strLogon_Time               As String
    Dim strCurrent_User_Name        As String
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        Call Get_Current_Logon_User(strLogon_Time, strCurrent_User, strCurrent_User_Name)
        strQuery = "SELECT * FROM USER_DATA WHERE "
        strQuery = strQuery & "USER_ID = '" & strCurrent_User & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
'            With pUser_Data
'                .USER_ID = lstRecord.Fields("USER_ID")
'                .USER_NAME = lstRecord.Fields("USER_NAME")
'                .ID_CARD_CODE = lstRecord.Fields("ID_CARD_CODE")
'                .USER_PW1 = lstRecord.Fields("USER_PW1")
'                .USER_PW2 = lstRecord.Fields("USER_PW2")
                CURRENT_USER_LEVEL = lstRecord.Fields("USER_LEVEL")
'            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    Else
        Call SaveLog("Get_User_Level", strDB_Path & strDB_FileName & " does not found.")
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Get_User_Level", ErrMsg)
    dbMyDB.Close
    
End Sub

Public Function Get_Current_User_Level() As String

    Get_Current_User_Level = CURRENT_USER_LEVEL
    
End Function

Public Sub Set_Download_Flag(ByVal pFlag As String)

    If pFlag <> CONTROL_UPDATE_TIME Then
        CONTROL_UPDATE_TIME = pFlag
        DOWNLOAD_FLAG = "E"
'==========================================================================================================
'
'  Modify Date : 2011. 12. 20
'  Modify by K.H. KIM
'  Content
'    - If Control data update time is not diffrent DOWNLOAD_FLAG set to "D"
'
'
'  Start of modify
'
'==========================================================================================================
'    Else
'        DOWNLOAD_FLAG = "D"
'===========================================================================================================
'
'  End of modify
'
'===========================================================================================================
    End If
    
End Sub

Public Function Get_Download_Flag() As String

    Get_Download_Flag = DOWNLOAD_FLAG
    
End Function

Public Sub Reset_Download_Flag()

    DOWNLOAD_FLAG = "D"
    
End Sub

Public Sub Set_Data_Change(ByVal pChange As Boolean)

    DATA_CHANGE = pChange
    
End Sub

Public Function Get_Data_Change() As Boolean

    Get_Data_Change = DATA_CHANGE
    
End Function
