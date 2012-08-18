VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal dirPath As String) As Long





Public Function do_Upload(ByVal pFileName As String, ByVal pRemote_Folder As String, ByVal pLocal_Folder As String) As Boolean
    Dim fs As Object
    Dim sourceFile As String
    Dim destinationFile As String
    Dim remoteDisk As String
  
   
    remoteDisk = "\\" & get_Remote_Server_IP & "\" & get_Remote_Server_Folder & "\"
    
    sourceFile = pLocal_Folder + pFileName
    destinationFile = remoteDisk + pRemote_Folder + pFileName
    
    If Dir(remoteDisk + pRemote_Folder, vbDirectory) = "" Then MakeSureDirectoryPathExists remoteDisk + pRemote_Folder
    Set fs = CreateObject("Scripting.FileSystemObject")  '创建声明
    fs.CopyFile sourceFile, destinationFile, True
    do_Upload = True
'    fs.CopyFolder "d:\在", "\\192.168.1.90\QLL\新建文件夹", True
'                            '源文件夹     目标文件夹
'Shell "net   use   \\192.168.1.90   526456   /user:qll ", vbHide '当窗体加载的时候输入远程服务器的登陆用户名和秘码
End Function

Public Function Write_Remote_Index(ByVal pFileName As String, ByVal pRemote_Folder As String)
   
   Dim strMessage As String
    strMessage = "Test"
'    WriteText(pFileName,pLocal_Folder,strMessage)
    Write_Remote_Index = True
    'FileHandle   =   OpenFile( "\\ "   &   IP地址   &   文件路径   &   "\ "   &   文件名,   lpReOpenBuff,   OF_READ)

End Function

Public Function Write_Local_Index(ByVal pFileName As String, ByVal pLocal_Folder As String)
    Dim strMessage As String
    strMessage = "Test"
    Dim strTime As String
    
    Dim strDate As String
    Dim strFileName As String
    
    strTime = Format(TIME, "HH:MM:SS")
    strDate = Format(DATE, "YYYY-MM-DD")
    strFileName = Format(DATE, "YYYYMMDD") & Format(TIME, "HH") & ".csv"
    
'    WriteText(pFileName,pLocal_Folder,strMessage)
    Write_Local_Index = True
    
End Function
Private Function WriteText(ByVal pFileName As String, ByVal pFolder As String, ByVal pMESSAGE As String)
   
    
    Dim intFileNum                          As Integer
    Dim ErrMsg As String

    intFileNum = FreeFile
    On Error GoTo ErrorHandler
    Open pFolder & pFileName For Append As intFileNum
    
    Print #intFileNum, pMESSAGE
    
    Close intFileNum
    
'    Shell "net use \\机器名 密码 /user:用户名"
'或者shell "cmd/c explorer\\机器名\文件夹名", 0
'Open ".txt" For Append As #1
'Close #1

ErrorHandler:
    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Decode_CST_Information_Elements", ErrMsg)
    Call Show_Message("Write index error", ErrMsg)
End Function
Public Function Check_Network() As Boolean

  Dim networkChecker As New clsNet
  
  Check_Network = networkChecker.IPValid(get_Remote_Server_IP)
  
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

Private Function get_Remote_Server_IP() As String
get_Remote_Server_IP = get_conf("Remote Server IP")
End Function

Private Function get_Remote_Server_Folder() As String
get_Remote_Server_Folder = get_conf("Remote Server Folder")
End Function

Public Function Get_Upload_Mode()
    Dim Ftp_Mode As Boolean
    Dim strKeyValue  As String
    
    strKeyValue = get_conf("Use FTP Mode")
    Ftp_Mode = True
    If strKeyValue = "1" Then
        Ftp_Mode = True
    Else
        Ftp_Mode = False
    End If
                    
    Get_Upload_Mode = Ftp_Mode 'True:use FTP uploading mode, false:use common uploading mode
End Function
