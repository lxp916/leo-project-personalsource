VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileExchanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' class added by Leo 2012.08.19
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal dirPath As String) As Long
Private Declare Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long
Public Function Get_File_From_Remote(ByVal pFileName As String, ByVal pRemote_Path As String, ByVal pLocal_Path As String) As Boolean
    If Right(pRemote_Path, 1) <> "\" Then
        pRemote_Path = pRemote_Path & "\"
    End If
    If Right(pLocal_Path, 1) <> "\" Then
        pLocal_Path = pLocal_Path & "\"
    End If
    
On Error GoTo ErrorHandler
    ' download files in file list
    Call Get_Remote_FileList(pFileName, pRemote_Path)
    'download file
    Call do_DownLoad(pFileName, pRemote_Path, pLocal_Path)
    Get_File_From_Remote = True
Exit Function

ErrorHandler:

    Call SaveLog("Get_File_From_Remote", pFileName & " upload faild. Remote path : " & pRemote_Path)
    Get_File_From_Remote = False
End Function
Public Function Get_Remote_FileList(ByVal pFileName As String, ByVal pRemote_Path As String) As String
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
        pRemote_Path = "\\" & get_Remote_Server_IP & "\" & get_Remote_Server_RootFolder & pRemote_Path & pFileName
        pRemote_Path = Replace(pRemote_Path, "\\\", "\")
        
        typFile_Data.cFileName = String(MAX_PATH, 0)    'Initialize filename value in WIN32_FIND_DATA structure
            hFind = FindFirstFile(pRemote_Path, typFile_Data)    'Find first file in the current directory
            lngLastError = Err.LastDllError     'Get last DLL error

            If hFind = 0 Then
                If (lngLastError = ERROR_NO_MORE_FILES) Then        'Can not find and file in the current directory
                    Call SaveLog("Get_Remote_FileList", pFileName & " does not exist.")
                    Get_Remote_FileList = ""
                Else                                                'Other error
                    Call SaveLog("Get_Remote_FileList", lngLastError)
                    Get_Remote_FileList = ""
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
                    lngResult = FindNextFile(hFind, typFile_Data)   'Find next correct file in the current directory
                    If lngResult = 0 Then   'Error happen
                        lngError = Err.LastDllError
                        If lngError = ERROR_NO_MORE_FILES Then
                            bolEOF = True
                            Get_Remote_FileList = strFileName
                        Else
                            bolEOF = True
                            Get_Remote_FileList = ""
                            Call SaveLog("Get_Remote_FileList", lngLastError)
                        End If
                    Else
                        strData = Left(typFile_Data.cFileName, InStr(1, typFile_Data.cFileName, String(1, 0), vbBinaryCompare) - 1)
                        Print #intFileNum, strData
                        strData = ""
                    End If
                Wend
                Close intFileNum

            End If

        DoEvents

        Call InternetCloseHandle(hFind)

    End If
End Function
Public Function Get_Remote_File_From_List(ByVal pRemotePath As String, ByVal pLocalPath As String, ByVal pPath As String, ByVal pFileName As String) As Boolean
    Dim strRemote_Path                  As String
    Dim strFileName                         As String
    
    Dim intFileNum                      As Integer
    Dim intResult                       As Integer
    Dim intLoopCount                    As Integer
    
    Dim lngIndex                        As Long
    
    Dim bolFind                         As Boolean
    
    If Dir(pPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open pPath & pFileName For Input As intFileNum
        While Not EOF(intFileNum)
            Line Input #intFileNum, strFileName
            If strFileName <> "" Then
                If do_DownLoad(strFileName, pRemotePath, pLocalPath) Then
                    Get_Remote_File_From_List = True
                    Call SaveLog("Get_Remote_File_From_List", strFileName & " download success")
                Else
                    Get_Remote_File_From_List = False
                    Call SaveLog("Get_Remote_File_From_List", strFileName & " download fail...")
                End If
                
            End If
            DoEvents
        Wend
        
        Close intFileNum
    End If
End Function

Private Function do_DownLoad(ByVal pFileName As String, ByVal pRemote_Path As String, ByVal pLocal_Path As String) As Boolean
    Dim fs As Object
    Dim localFile As String
    Dim remoteFile As String
    Dim returnValue As Boolean
    
On Error GoTo ErrorHandler
    If Check_Network = True Then
        remoteFile = "\\" & get_Remote_Server_IP & "\" & get_Remote_Server_RootFolder & pRemote_Path & pFileName
        remoteFile = Replace(remoteFile, "\\\", "\")
        
        localFile = pLocal_Path & pFileName
        
        If Dir(pLocal_Path, vbDirectory) = "" Then MakeSureDirectoryPathExists pLocal_Path
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Dir(remoteFile, vbNormal) <> "" Then
            fs.CopyFile remoteFile, localFile, True
            returnValue = True
        Else
            returnValue = False
        End If
    End If
    do_DownLoad = returnValue
    
    Exit Function
    
ErrorHandler:

    Call SaveLog("do_DownLoad", pFileName & " download faild. Remote path : " & pRemote_Path)
    do_DownLoad = False
End Function


Public Function do_Upload(ByVal pFileName As String, ByVal pLocal_Folder As String, ByVal pRemote_Folder As String) As Boolean
    Dim fs As Object
    Dim sourceFile As String
    Dim destinationFile As String
    Dim remoteDisk As String
    Dim returnValue As Boolean
On Error GoTo ErrorHandler

    remoteDisk = "\\" & get_Remote_Server_IP & "\" & get_Remote_Server_RootFolder & "\"
    
    sourceFile = pLocal_Folder & pFileName
    destinationFile = remoteDisk & pRemote_Folder & pFileName
    
    If Dir(remoteDisk & pRemote_Folder, vbDirectory) = "" Then MakeSureDirectoryPathExists remoteDisk & pRemote_Folder
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Dir(sourceFile, vbNormal) <> "" Then
        fs.CopyFile sourceFile, destinationFile, True
        returnValue = True
    Else
        returnValue = False
    End If
    do_Upload = returnValue
    
      Exit Function
    
ErrorHandler:

    Call SaveLog("do_Upload", pFileName & " upload faild. Remote path : " & pRemote_Folder)
    do_Upload = False
End Function

Public Function do_upload_files() As Boolean
    Dim dbMyDB                              As Database
    Dim lstRecord                           As Recordset
    Dim intCount                            As Integer
    Dim strQuery                            As String
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strFileName As String
    Dim strlocalFolder As String
    Dim strRemoteFolder As String
    Dim strRecordID As String
    Dim isUpload As Boolean

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        strQuery = "SELECT * FROM LOCAL_INDEX"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        strRecordID = ""
        isUpload = False
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                strFileName = lstRecord.Fields("file_name")
                strlocalFolder = lstRecord.Fields("local_path")
                strRemoteFolder = lstRecord.Fields("remote_path")
                
                If Dir(strlocalFolder & strFileName, vbNormal) <> "" Then
                    isUpload = do_Upload(strFileName, strlocalFolder, strRemoteFolder)
                    If isUpload Then
                        Call Write_Remote_Index(strFileName, strRemoteFolder)
                        strRecordID = strRecordID & lstRecord.Fields("ID") & ","
                    End If
                End If
                
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        If (Len(strRecordID) > 0) Then
            strRecordID = "(" & strRecordID & "0" & ")"
            strQuery = "Delete FROM LOCAL_INDEX WHERE ID IN" & strRecordID
            dbMyDB.Execute (strQuery)
        End If
        dbMyDB.Close
    End If
    do_upload_files = True

End Function

Public Function hasExistingLocalIndex() As String
    Dim dbMyDB                              As Database
    Dim lstRecord                           As Recordset
    Dim strDB_Path As String
    Dim strDB_FileName As String
    Dim strQuery                            As String
 
    Dim hasRecord As Boolean

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        strQuery = "SELECT * FROM LOCAL_INDEX"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        If lstRecord.RecordCount = 0 Then
            hasRecord = False
        Else: hasRecord = True
        End If
        
        dbMyDB.Close
    End If
    hasExistingLocalIndex = hasRecord
End Function

Public Function Write_Remote_Index(ByVal pFileName As String, ByVal pRemote_Folder As String)
    Dim strMessage As String
    Dim strPath  As String
    Dim strType As String
    Dim remoteFolder() As String
    
    remoteFolder = Split(pRemote_Folder, "\")
    
    strPath = "\\" & get_Remote_Server_IP & "\" & get_Remote_Server_RootFolder & "Index\"
    If UBound(remoteFolder) > 0 Then
        strType = remoteFolder(0)
        strPath = strPath & strType & "\"
    End If
    
    strMessage = Format(DATE, "YYYY/MM/DD") & " " & Format(TIME, "HH:MM:SS") & "," & pRemote_Folder & pFileName
    
    Call WriteIndex(generate_IndexFileName, strPath, strMessage)
    Write_Remote_Index = True

End Function

Public Function Write_Local_Index(ByVal pFileName As String, ByVal pLocal_Folder As String, ByVal pRemote_Folder As String)
    Dim strMessage As String
    Dim strDB_Path                      As String
    Dim strDB_FileName                  As String
    Dim strQuery As String
    Dim dbMyDB                              As Database
    
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"
     If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "INSERT INTO LOCAL_INDEX (date_time,file_name,local_path,remote_path) Values ("
        strQuery = strQuery & "'" & Format(DATE, "YYYY/MM/DD") & " " & Format(TIME, "HH:MM:SS") & "',"
        strQuery = strQuery & "'" & pFileName & "',"
        strQuery = strQuery & "'" & pLocal_Folder & "',"
        strQuery = strQuery & "'" & pRemote_Folder & "'"
        strQuery = strQuery & ")"
        dbMyDB.Execute (strQuery)
        dbMyDB.Close
        Write_Local_Index = True
    Else
        Write_Local_Index = False
    End If
        
'    strMessage = Format(DATE, "YYYY/MM/DD") & " " & Format(TIME, "HH:MM:SS") & "," & pFileName & "," & pLocal_Folder & "," & pRemote_Folder
   
    
'    Call WriteIndex(generate_IndexFileName, get_LocalFailIndexFolder, strMessage)
    Write_Local_Index = True

End Function
Private Function WriteIndex(ByVal pFileName As String, ByVal pFolder As String, ByVal pMESSAGE As String)
   
    
    Dim intFileNum                          As Integer
    Dim fs As Object
    Dim ErrMsg As String

    Set fs = CreateObject("Scripting.FileSystemObject")
    intFileNum = FreeFile
    
    If Dir(pFolder & pFileName, vbDirectory) = "" Then
        MakeSureDirectoryPathExists pFolder & pFileName
        
        Open pFolder & pFileName For Output As intFileNum
        Print #intFileNum, "date_time,data_file"
        
        Close intFileNum
    End If
    
    Open pFolder & pFileName For Append As intFileNum
    
    Print #intFileNum, pMESSAGE
    
    Close intFileNum

End Function
Public Function Check_Network() As Boolean

  Dim networkChecker As New clsNet
  Dim isConnected As Boolean
  Dim server_IP As String
  Dim login_id As String
  Dim pwd As String
  Dim strShellCMD As String
  
  server_IP = get_Remote_Server_IP
  login_id = get_remote_server_login
  pwd = get_remote_server_pwd
  strShellCMD = "net use \\" & server_IP & " " & pwd & " / USER:" & login_id
  
  isConnected = networkChecker.IPValid(server_IP)
  If isConnected = False Then
    Shell strShellCMD
    isConnected = networkChecker.IPValid(server_IP)
  End If
  
  ' retry 3 times after every 10 seconds if network is not reachable
  
  If isConnected = False Then
    Sleep 10000
    Shell strShellCMD
    isConnected = networkChecker.IPValid(server_IP)
  End If
  
    
  If isConnected = False Then
    Sleep 10000
    Shell strShellCMD
    isConnected = networkChecker.IPValid(server_IP)
  End If
  
   
  If isConnected = False Then
    Sleep 10000
    Shell strShellCMD
    isConnected = networkChecker.IPValid(server_IP)
  End If
  
  Check_Network = isConnected
  
End Function

Private Function generate_IndexFileName() As String
    generate_IndexFileName = Format(DATE, "YYYYMMDD") & ".csv"
End Function

Private Function get_remote_server_login() As String
    get_remote_server_login = get_conf("Remote Server Login")
End Function

Private Function get_remote_server_pwd() As String
    get_remote_server_pwd = get_conf("Remote Server Login Password")
End Function

Private Function get_Remote_Server_IP() As String
get_Remote_Server_IP = get_conf("Remote Server IP")
End Function

Private Function get_Remote_Server_RootFolder() As String
    Dim pRemotePath As String
    pRemotePath = get_conf("Remote Server Folder")
    If Right(pRemotePath, 1) <> "\" Then
        pRemotePath = pRemotePath & "\"
    End If
    get_Remote_Server_RootFolder = pRemotePath
End Function

Public Function get_Remote_Server_ReadOnlyFolder() As String
    Dim pRemotePath As String
'    pRemotePath = get_conf("HOST DATA PATH")
    pRemotePath = "EQ_Config\JPS\Readonly\"
    If Right(pRemotePath, 1) <> "\" Then
        pRemotePath = pRemotePath & "\"
    End If
    get_Remote_Server_ReadOnlyFolder = pRemotePath
End Function


Private Function get_LocalFailIndexFolder() As String
get_LocalFailIndexFolder = App.PATH & "\DB\Fail_Index\"
End Function

Public Function IsFTPUploadMode()
    Dim Ftp_Mode As Boolean
    Dim strKeyValue  As String
    
    strKeyValue = get_conf("Use FTP Mode")
    Ftp_Mode = True
    If strKeyValue = "1" Then
        Ftp_Mode = True
    Else
        Ftp_Mode = False
    End If
                    
    IsFTPUploadMode = Ftp_Mode 'True:use FTP uploading mode, false:use common uploading mode
End Function

Private Function get_conf(ByVal strKey As String) As String
    Dim strKeyValue As String
    Dim strFileName As String
    Dim intFileNum As Integer
    Dim strTemp As String
    Dim intPos As Integer
    Dim intPosition As Integer

    strFileName = App.PATH & "\Env\FTP_Parameter.cfg"
    If Dir(strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                Select Case Left(strTemp, intPos - 1)
                Case strKey:
                    strKeyValue = Mid(strTemp, intPos + 1)
                End Select
            End If
        Wend
    End If
    Close intFileNum
    get_conf = strKeyValue
End Function
