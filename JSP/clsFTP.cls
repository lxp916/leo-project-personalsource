VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim bActiveSession              As Boolean

Dim IP_ADDRESS                  As String
Dim USER_ID                     As String
Dim PASSWORD                    As String
Dim INDEX_PATH                  As String
Dim DEFECT_PATH                 As String
Dim IMAGE_PATH                  As String

Dim USE_FTP                     As Boolean

Dim PORT_NUMBER                 As Long
Dim hOpen                       As Long
Dim hConnection                 As Long
Dim dwType                      As Long

Public Function Init_FTP_Client() As Boolean

    Dim intFileNum          As Integer
    
    Dim strFileName         As String
    Dim strTemp             As String
    
    Dim intPosition         As Integer
    
    USE_FTP = False
    strFileName = App.PATH & "\Env\FTP_Parameter.cfg"
    If Dir(strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strFileName For Input As intFileNum
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPosition = InStr(strTemp, "=")
            Select Case Left(strTemp, intPosition - 1)
            Case "IP ADDRESS":
                IP_ADDRESS = Mid(strTemp, intPosition + 1)
            Case "PORT NUMBER":
                PORT_NUMBER = CLng(Mid(strTemp, intPosition + 1))
            Case "USERID":
                USER_ID = Mid(strTemp, intPosition + 1)
            Case "PASSWORD":
                PASSWORD = Mid(strTemp, intPosition + 1)
            Case "HOST DATA PATH":
                INDEX_PATH = Mid(strTemp, intPosition + 1)
            Case "DEFECT PATH":
                DEFECT_PATH = Mid(strTemp, intPosition + 1)
            Case "USE FTP":
                If Mid(strTemp, intPosition + 1) = "1" Then
                    USE_FTP = True
                Else
                    USE_FTP = False
                End If
            End Select
        Wend
        Close intFileNum
        
        If USE_FTP = True Then
            If (IP_ADDRESS <> "") And (USER_ID <> "") And (PASSWORD <> "") Then
                Init_FTP_Client = True
            Else
                Init_FTP_Client = False
            End If
        Else
            Init_FTP_Client = False
        End If
    Else
        IP_ADDRESS = ""
        USER_ID = ""
        PORT_NUMBER = 0
        PASSWORD = ""
        INDEX_PATH = ""
        DEFECT_PATH = ""
        IMAGE_PATH = ""
        Init_FTP_Client = False
    End If

End Function

Public Sub Set_ActiveSession(ByVal pSession As Boolean)

    bActiveSession = pSession
    
End Sub

Public Function Get_ActiveSession() As Boolean

    Get_ActiveSession = bActiveSession
    
End Function

Public Sub Set_Open_Handle(ByVal pState As Long)

    hOpen = pState
    
End Sub

Public Function Get_Open_Handle() As Long

    Get_Open_Handle = hOpen
    
End Function

Public Sub Set_Connect_Handle(ByVal pHandle As Long)

    hConnection = pHandle
    
End Sub

Public Function Get_Connect_Handle() As Long

    Get_Connect_Handle = hConnection
    
End Function

Public Sub Set_Connect_Type(ByVal pType As Long)

    dwType = pType
    
End Sub

Public Function Get_Connect_Type() As Long

    Get_Connect_Type = dwType
    
End Function

Public Sub Open_Session()

    bActiveSession = False
    hOpen = 0
    hConnection = 0
    dwType = FTP_TRANSFER_TYPE_BINARY
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hOpen <> 0 Then
        Call SaveLog("Open_Session", "FTP Port Open")
        Call Connect_FTP_Client
    Else
        Call SaveLog("Open_Session", "FTP Port Open Error")
    End If

End Sub

Public Sub Connect_FTP_Client()

    Dim nFlag           As Long
    Dim lngError_Code   As Long
    
    If (bActiveSession = False) And (hOpen <> 0) Then
        nFlag = 0 'INTERNET_FLAG_PASSIVE
'        hConnection = InternetConnect(hOpen, ENV.Get_FTP_Server, INTERNET_DEFAULT_FTP_PORT, _
                                      USERID, PASSWORD, INTERNET_SERVICE_FTP, nFlag, 0)
        hConnection = InternetConnect(hOpen, IP_ADDRESS, PORT_NUMBER, _
                                      USER_ID, PASSWORD, INTERNET_SERVICE_FTP, nFlag, 0)
        If hConnection = 0 Then
            bActiveSession = False
            lngError_Code = GetLastError
            Call SaveLog("Connect_FTP_Client", "FTP Connect Fail")
        Else
            bActiveSession = True
            Call SaveLog("Connect_FTP_Client", "FTP Connect Succeed")
            Call Remote_Change_Directory("\")
       End If
    End If

End Sub

Public Sub Disconnect_FTP_Client()

    If hConnection <> 0 Then
        InternetCloseHandle hConnection     'Disconnect FTP connection
    End If
    hConnection = 0
    bActiveSession = False
    Call SaveLog("Disconnect_FTP_Client", "FTP Disconnect")
    
End Sub

Public Sub Close_Session()

    If hConnection <> 0 Then
        InternetCloseHandle (hConnection)   'Disconnect FTP connection
    End If
    If hOpen <> 0 Then
        InternetCloseHandle (hOpen)         'Close Session
    End If
    hConnection = 0
    hOpen = 0
    bActiveSession = False

End Sub

Public Function Remote_Change_Directory(pDir As String) As Integer  '0:Directory change OK, 1:Directory not exist, 2:FTP Connection fail, else:Error code
    
    Dim strPathFromRoot     As String
    Dim strTemp_Path        As String
    Dim strCurrent_Path     As String
    
    Dim intPosition         As Integer
    
    Dim hResp               As Long
    
    Dim bolResult           As Boolean
    
'    Dim Fso                 As Object
    
    Dim ErrMsg              As String
    
On Error GoTo ErrorHandler

    Remote_Change_Directory = 0
    If pDir = "" Then
        Call SaveLog("Remote_Change_Directory", "Directory does not exist. Folder Name : " & pDir)
    Else
        If InStr(1, pDir, IP_ADDRESS) Then
            strPathFromRoot = Mid(pDir, Len(IP_ADDRESS) + 1, Len(pDir) - Len(IP_ADDRESS))
        Else
            strPathFromRoot = pDir
        End If
        If strPathFromRoot = "" Then
            strPathFromRoot = "\"
        End If
        If Right(strPathFromRoot, 1) <> "\" Then
            strPathFromRoot = strPathFromRoot & "\"
        End If
        Call SaveLog("Remote_Change_Directory", "strPathFromRoot : " & strPathFromRoot)
        bolResult = FtpSetCurrentDirectory(hConnection, "/Cell")
        strCurrent_Path = Get_Current_Directory
        If strCurrent_Path = "" Then
            Remote_Change_Directory = 2
'            Call Disconnect_FTP_Client
'            Call Close_Session
'            If Init_FTP_Client = True Then
'                Call FTP_CLIENT.Open_Session
'            Else
'                Call SaveLog("Remote_Change_Directory", "FTP Client initialize fail.")
'            End If
        Else
            strTemp_Path = strPathFromRoot
'            strCurrent_Path = ""
            If Left(strTemp_Path, 1) <> "\" Then
                strTemp_Path = "/" & Trim(strTemp_Path)
            End If
            intPosition = InStr(strTemp_Path, "\")
            While intPosition <> 0
                strCurrent_Path = strCurrent_Path & Trim(Left(strTemp_Path, intPosition - 1))
                strCurrent_Path = strCurrent_Path & "/"
                strTemp_Path = Mid(strTemp_Path, intPosition + 1)
                If FtpSetCurrentDirectory(hConnection, strCurrent_Path) = False Then
'                    Set Fso = CreateObject("Scripting.FileSystemObject")
'                    Fso.Attributes = Fso.Attributes + 32
                    If Create_Remote_Directory(strCurrent_Path) = False Then
                        Call SaveLog("Remote_Change_Direcotry", strCurrent_Path & " create fail.")
                        Remote_Change_Directory = 2
                        intPosition = 0
                    Else
'                        If FtpCommand(hConnection, True, FTP_TRANSFER_TYPE_BINARY, "SITE chmod 777 " & strCurrent_Path, 0, hResp) = True Then
                            If FtpSetCurrentDirectory(hConnection, strCurrent_Path) = False Then
                                Call SaveLog("Remote_Change_Directory", strCurrent_Path & " change fail.")
                                Remote_Change_Directory = 2
                                intPosition = 0
                            Else
                                If strTemp_Path <> "" Then
                                    intPosition = InStr(strTemp_Path, "\")
                                Else
                                    Remote_Change_Directory = 0
                                    Call SaveLog("Remote_Change_Directory", "Directory change success. Current directory : " & strCurrent_Path)
                                    intPosition = 0
                                End If
                            End If
'                        Else
'                            Call SaveLog("Remote_Change_Directory", "Mode change fail.")
'                            Remote_Change_Directory = 2
'                            intPosition = 0
'                        End If
                    End If
                Else
'                    Call SaveLog("Remote_Change_Directory", "Directory change success. Current directory : " & strCurrent_Path)
                    If strTemp_Path <> "" Then
                        intPosition = InStr(strTemp_Path, "\")
                    Else
                        Remote_Change_Directory = 0
                        Call SaveLog("Remote_Change_Directory", "Directory change success. Current directory : " & strCurrent_Path)
                        intPosition = 0
                    End If
                End If
            Wend
        End If
    End If
    
    Exit Function
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Remote_Change_Direcotry", ErrMsg)
    Remote_Change_Directory = Err.Number
    
End Function

Public Function SetFTPDirectory(sDir As String)
    
    If (FtpSetCurrentDirectory(hConnection, sDir) = False) Then
        SetFTPDirectory = False
    Else
        SetFTPDirectory = True
    End If

End Function

Public Function GetFTPDirectory() As String
    
    Dim szDir           As String
    
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnection, szDir, 1024) = False) Then
        GetFTPDirectory = ""
    Else
        GetFTPDirectory = Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If

End Function

Public Function Get_Current_Directory() As String

    Dim strPath         As String
    
    strPath = String(1024, Chr$(0))
    
    If (FtpGetCurrentDirectory(hConnection, strPath, 1024) = False) Then
        Call SaveLog("Get_Current_Directory", "FtpGetCurrentDirectory funciton execute failed.")
    Else
        Get_Current_Directory = Left(strPath, InStr(1, strPath, String(1, 0), vbBinaryCompare) - 1)
        Call SaveLog("Get_Current_Directory", "Current Directory : " & Get_Current_Directory)
    End If
    
End Function

Public Function Create_Remote_Directory(ByVal pPath As String) As Boolean

    If (FtpCreateDirectory(hConnection, pPath) = False) Then
        Call SaveLog("Create_Remote_Directory", pPath & " create failed.")
        Create_Remote_Directory = False
    Else
        Call SaveLog("Create_Remote_Directory", pPath & " create succeed.")
        Create_Remote_Directory = True
    End If

End Function

Public Function FTP_Put_File(ByVal pFileName As String, ByVal pRemote_Folder As String, ByVal pLocal_Folder As String) As Boolean

    Dim bolResult           As Boolean
    
    Dim strFileRemote       As String
    Dim strFileLocal        As String
    Dim strDirRemote        As String
    Dim strDirLocal         As String
    Dim strDirTemp          As String
    
    Dim intPosition         As Integer
    Dim intResult           As Integer
    Dim intReconnect_Count  As Integer
    
    Dim Fso                 As Object
    Dim objFso
    
    Dim lngIndex            As Long
    
    If bActiveSession = True Then
        If Right(pRemote_Folder, 1) <> "\" Then
            pRemote_Folder = pRemote_Folder & "\"
        End If
        
        If Right(pLocal_Folder, 1) <> "\" Then
            pLocal_Folder = pLocal_Folder & "\"
        End If
        
        strFileRemote = pFileName
        strFileLocal = pFileName
        
        intReconnect_Count = 0
        While intReconnect_Count < 3
            intReconnect_Count = intReconnect_Count + 1
            intResult = Remote_Change_Directory(pRemote_Folder)
            Select Case intResult
            Case 0:     'Current directory change complete
'                Set Fso = CreateObject("Scripting.FileSystemObject")
'                Set objFso = Fso.getfile(pLocal_Folder & strFileLocal)
'                objFso.Attributes = Fso.Attributes + 32
                bolResult = FtpPutFile(hConnection, pLocal_Folder & strFileLocal, strFileRemote, dwType, 0)
                If bolResult = False Then
                    Call SaveLog("FTP_Put_File", "FTP File Write fail. File Name : " & strFileLocal)
                    intReconnect_Count = 3
                    FTP_Put_File = False
                Else
                    Call SaveLog("FTP_PUT_File", "FTP File Write succeed. File Name : " & strFileLocal)
                    intReconnect_Count = 3
                    FTP_Put_File = True
                End If
            Case 1:     'Directory does not exist
    '                If FTP_Create_Directory(strDirRemote) = True Then
    '                    intResult = Remote_Change_Directory(strDirRemote)
    '                    If intPosition < Len(strDirTemp) Then
    '                        strDirTemp = Mid(strDirTemp, intPosition + 1)
    '                    Else
    '                        strDirTemp = ""
    '                        bolChange_Finish = False
    '                    End If
    '                    intPosition = InStr(strDirTemp, "\")
    '                Else
    '                    intPosition = 0
    '                    bolChange_Finish = False
    '                End If
            Case 2:     'FTP Connection fail
                Call Disconnect_FTP_Client
                Call Close_Session
                Call SaveLog("FTP_PUT_File", "Remote directory chage fail.")
'                If Init_FTP_Client = True Then
'                    Call FTP_CLIENT.Open_Session
'                    intReconnect_Count = intReconnect_Count + 1
'                    Call SaveLog("FTP_PUT_File", "Reconnect count : " & intReconnect_Count)
'                Else
'                    intPosition = 0
'                    intReconnect_Count = 3
'                    Call SaveLog("Remote_Change_Directory", "FTP Client initialize fail.")
'                End If
            Case Else
                intReconnect_Count = 3
            End Select
            DoEvents
        Wend
    Else
        FTP_Put_File = False
   End If

End Function

Public Function FTP_Get_File(ByVal pFileName As String, ByVal pRemote_Path As String, ByVal pLocal_Path As String) As Boolean

    Dim strFileRemote       As String
    Dim strDirRemote        As String
    Dim strFileLocal        As String
    Dim strFileName         As String
    
    Dim bolResult           As Boolean
    Dim bolFind             As Boolean
    
    Dim lngIndex            As Long
    Dim lngError_Code       As Long
    
    Dim intLoopCount        As Integer
    Dim intResult           As Integer
    
    If bActiveSession = True Then
        If Right(pRemote_Path, 1) <> "\" Then
            pRemote_Path = pRemote_Path & "\"
        End If
        If Right(pLocal_Path, 1) <> "\" Then
            pLocal_Path = pLocal_Path & "\"
        End If
        
        strFileName = FTP_Get_FileList(pFileName, pRemote_Path)
'        If strFileName <> "" Then
            intResult = Remote_Change_Directory(pRemote_Path)
            Select Case intResult
            Case 0:
                intResult = FtpGetFile(hConnection, pFileName, pLocal_Path & pFileName, False, FILE_ATTRIBUTE_NORMAL, dwType, 0)
                If intResult = 1 Then
                    bolFind = False
                    intLoopCount = 0
                    For lngIndex = 1 To 10000
                        DoEvents
                    Next lngIndex
                    While bolFind = False
                        If Dir(pLocal_Path & pFileName, vbNormal) <> "" Then
                            bolFind = True
                        Else
                            If intLoopCount < 5 Then
                                For lngIndex = 1 To 10000
                                    DoEvents
                                Next lngIndex
                                intLoopCount = intLoopCount + 1
                            Else
                                bolFind = True
                            End If
                        End If
                    Wend
                    If intLoopCount >= 5 Then
                        FTP_Get_File = False
                    Else
                        FTP_Get_File = True
                    End If
                Else
                    lngError_Code = GetLastError
                    FTP_Get_File = False
                End If
            Case 1:
                Call SaveLog("FTP_Get_File", "Sub directory is not exist.")
                FTP_Get_File = False
            Case 2:
                Call SaveLog("FTP_Get_File", "FTP Connectin fail.")
                FTP_Get_File = False
            Case Else
                Call SaveLog("FTP_Get_File", "Error Code : " & intResult)
                FTP_Get_File = False
            End Select
            
            DoEvents
'        Else
'            FTP_Get_File = False
'        End If
    Else
        FTP_Get_File = False
    End If
    
End Function

Public Function FTP_Get_File_from_List(ByVal pRemotePath As String, ByVal pLocalPath As String, ByVal pPath As String, ByVal pFileName As String) As Boolean

    Dim strRemote_Path                  As String
    Dim strTemp                         As String
    
    Dim intFileNum                      As Integer
    Dim intResult                       As Integer
    Dim intLoopCount                    As Integer
    
    Dim lngIndex                        As Long
    
    Dim bolFind                         As Boolean
    
    If Dir(pPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        If Init_FTP_Client = True Then
            Call Open_Session
            If Right(pRemotePath, 1) <> "\" Then
                pRemotePath = pRemotePath & "\"
            End If
'            pRemotePath = pRemotePath & "Table\"
            If bActiveSession = True Then
                If Right(pLocalPath, 1) <> "\" Then
                    pLocalPath = pLocalPath & "\"
                End If
                intResult = Remote_Change_Directory(pRemotePath)
                Select Case intResult
                Case 0:
                    Open pPath & pFileName For Input As intFileNum
                    
                    While Not EOF(intFileNum)
                        Line Input #intFileNum, strTemp
                        
                        If strTemp <> "" Then
                            intResult = FtpGetFile(hConnection, strTemp, pLocalPath & strTemp, False, FILE_ATTRIBUTE_NORMAL, dwType, 0)
                            If intResult = 1 Then
                                bolFind = False
                                intLoopCount = 0
                                While bolFind = False
                                    If Dir(pLocalPath & strTemp, vbNormal) <> "" Then
                                        bolFind = True
                                    Else
                                        If intLoopCount < 5 Then
                                            For lngIndex = 1 To 10000
                                                DoEvents
                                            Next lngIndex
                                            intLoopCount = intLoopCount + 1
                                        Else
                                            bolFind = True
                                        End If
                                    End If
                                Wend
                                If intLoopCount >= 5 Then
                                    FTP_Get_File_from_List = False
                                    Call SaveLog("FTP_Get_File_from_List", strTemp & " download fail..")
                                Else
                                    FTP_Get_File_from_List = True
                                    Call SaveLog("FTP_Get_File_from_List", strTemp & " download success.")
                                End If
                            End If
                        End If
                        DoEvents
                    Wend
                    
                    Close intFileNum
                Case 1:
                    Call SaveLog("FTP_Get_File_from_List", "Sub directory is not exist.")
                    FTP_Get_File_from_List = False
                Case 2:
                    Call SaveLog("FTP_Get_File_from_List", "FTP Connectin fail.")
                    FTP_Get_File_from_List = False
                Case Else
                    Call SaveLog("FTP_Get_File_from_List", "Error Code : " & intResult)
                    FTP_Get_File_from_List = False
                End Select
                
                DoEvents
            End If
            Call Close_Session
            Call Disconnect_FTP_Client
        Else
            Call SaveLog("FTP_Get_File_From_List", "FTP initialize fail.")
        End If
    End If
    
End Function

Public Function FTP_Delete_File(ByVal pFileName As String, ByVal pRemote_Path As String) As Boolean

    Dim strFileRemote       As String
    Dim strDirRemote        As String
    
    Dim bolResult           As Boolean
    
    If bActiveSession = True Then
        bolResult = FtpDeleteFile(hConnection, pRemote_Path & pFileName)
        If bolResult = True Then
            FTP_Delete_File = True
        Else
            FTP_Delete_File = False
        End If
    End If
    
End Function

Public Function Get_IP_Address() As String

    Get_IP_Address = IP_ADDRESS
    
End Function

Public Function Get_Port_Number() As String

    Get_Port_Number = PORT_NUMBER
    
End Function

Public Function Get_UserID() As String

    Get_UserID = USER_ID
    
End Function

Public Function Get_Password() As String

    Get_Password = PASSWORD
    
End Function

Public Function Get_Path(ByVal pType As Integer) As String

    Select Case pType
    Case cFTP_HOST:
        Get_Path = INDEX_PATH
    Case cFTP_DEFECT:
        Get_Path = DEFECT_PATH
    End Select
    
End Function

Public Function FTP_Create_Directory(ByVal pDirectory As String) As Boolean

    If Right(pDirectory, 1) <> "\" Then
        pDirectory = pDirectory & "\"
    End If
    
    If (FtpCreateDirectory(hConnection, pDirectory) = False) Then
        Call SaveLog("FTP_Crate_Directory", pDirectory & " create fail.")
        Call MsgBox(pDirectory & " create fail.", vbOKOnly, "FTP error")
        FTP_Create_Directory = False
    Else
        FTP_Create_Directory = True
    End If
    
End Function


Public Function FTP_Get_FileList(ByVal pFileName As String, ByVal pRemote_Path As String) As String
    
    Dim strFileName         As String
    Dim strData             As String
    Dim strPath             As String
    
    Dim hFind               As Long
    Dim lngLastError        As Long
    Dim lngError            As Long
    Dim lngResult           As Long
    
    Dim intResult           As Integer
    Dim intFileNum          As Integer
    
    Dim bolEOF              As Boolean
    
    Dim typFile_Data        As WIN32_FIND_DATA
    
    If Len(pRemote_Path) > 0 Then
        If Right(pRemote_Path, 1) <> "\" Then
            pRemote_Path = pRemote_Path & "\"
        End If
        intResult = Remote_Change_Directory(pRemote_Path)   'Change remote directory
        Select Case intResult
        Case 0:
            typFile_Data.cFileName = String(MAX_PATH, 0)    'Initialize filename value in WIN32_FIND_DATA structure
            hFind = FtpFindFirstFile(hConnection, pFileName, typFile_Data, 0, 0)      'Find first file in the current directory
            lngLastError = Err.LastDllError     'Get last DLL error
            
            If hFind = 0 Then
                If (lngLastError = ERROR_NO_MORE_FILES) Then        'Can not find and file in the current directory
                    Call SaveLog("FTP_Get_FileList", pFileName & " does not exist.")
                    FTP_Get_FileList = ""
                Else                                                'Other error
                    Call SaveLog("FTP_Get_FileList", lngLastError)
                    FTP_Get_FileList = ""
                End If
            Else
                lngError = NO_ERROR
                strPath = App.PATH & "\Env\"
                strFileName = "FTP_File_List.txt"
                intFileNum = FreeFile
                Open strPath & strFileName For Output As intFileNum
                
                'Get filename in WIN32_FIND_DATA structure
                strData = Left(typFile_Data.cFileName, InStr(1, typFile_Data.cFileName, String(1, 0), vbBinaryCompare) - 1)
                Print #intFileNum, strData      'Print out in the result file
                strData = ""
                
                bolEOF = False
                While bolEOF = False
                    typFile_Data.cFileName = String(MAX_PATH, 0)    'Initialize filename value in WIN32_FIND_DATA structure
                    lngResult = InternetFindNextFile(hFind, typFile_Data)   'Find next JPG file in the current directory
                    If lngResult = 0 Then   'Error happen
                        lngError = Err.LastDllError
                        If lngError = ERROR_NO_MORE_FILES Then
                            bolEOF = True
                            FTP_Get_FileList = strFileName
                        Else
                            bolEOF = True
                            FTP_Get_FileList = ""
                            Call SaveLog("FTP_Get_FileList", lngLastError)
                        End If
                    Else
                        strData = Left(typFile_Data.cFileName, InStr(1, typFile_Data.cFileName, String(1, 0), vbBinaryCompare) - 1)
                        Print #intFileNum, strData
                        strData = ""
                    End If
                Wend
                Close intFileNum
                
            End If
        Case 1:
            Call SaveLog("FTP_Get_FileList", "Sub directory is not exist.")
            FTP_Get_FileList = ""
        Case 2:
            Call SaveLog("FTP_Get_FileList", "FTP Connectin fail.")
            FTP_Get_FileList = ""
        Case Else
            Call SaveLog("FTP_Get_FileList", "Error Code : " & intResult)
            FTP_Get_FileList = ""
        End Select
        
        DoEvents
        
        Call InternetCloseHandle(hFind)
        
    End If
    
End Function

