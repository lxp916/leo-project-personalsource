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
Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
'常数和变量的设定
Const RAS95_MaxEntryName = 256
Const RAS95_MaxDeviceType = 16
Const RAS95_MaxDeviceName = 32

Private Type RASCONN95
  dwSize As Long
  hRasCon As Long
  szEntryName(RAS95_MaxEntryName) As Byte
  szDeviceType(RAS95_MaxDeviceType) As Byte
  szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private Type RASCONNSTATUS95
  dwSize As Long
  RasConnState As Long
  dwError As Long
  szDeviceType(RAS95_MaxDeviceType) As Byte
  szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Function Init_Upload() As Boolean
    Dim remoteDirectory As String
    
    Init_Upload = True
    
End Function

Public Function Get_Upload_Mode()
    Dim Ftp_Mode As Boolean
    Dim strFileName As String
    Dim intFileNum As Integer
    Dim strTemp As String
    Dim intPos As Integer
    Dim intPosition As Integer
    
    
    
    
    Ftp_Mode = True
    strFileName = App.PATH & "\Env\FTP_Parameter.cfg"
    If Dir(strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                Select Case Left(strTemp, intPos - 1)
                
                'added by leo
                Case "Use FTP Mode":
                    If Mid(strTemp, intPosition + 1) = "1" Then
                        Ftp_Mode = True
                    Else
                        Ftp_Mode = False
                    End If
                End Select
            End If
        Wend
    End If
        Close intFileNum
    Get_Upload_Mode = Ftp_Mode 'True:use FTP uploading mode, false:use common uploading mode
End Function

Public Function do_Upload(ByVal pFileName As String, ByVal pRemote_Folder As String, ByVal pLocal_Folder As String) As Boolean
    Dim fs As Object
    Dim sourceFile As String
    Dim destinationFile As String
    Dim remoteDisk As String
    
    remoteDisk = "F:\"
    
    sourceFile = pLocal_Folder + pFileName
    destinationFile = remoteDisk + pRemote_Folder + pFileName
    
    If Dir(remoteDisk + pRemote_Folder, vbDirectory) = "" Then MakeSureDirectoryPathExists remoteDisk + pRemote_Folder
    Set fs = CreateObject("Scripting.FileSystemObject")  '创建声明
    fs.CopyFolder sourceFile, destinationFile, True
    do_Upload = True
    
End Function

Public Function Write_Remote_Index(ByVal pFileName As String, ByVal pRemote_Folder As String)
   
   Dim strMessage As String
    strMessage = "Test"
'    WriteText(pFileName,pLocal_Folder,strMessage)
    Write_Remote_Index = True
    

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
Private Function WriteText(ByVal pFileName As String, ByVal pFolder As String, ByVal pMessage As String)
   
    
    Dim intFileNum                          As Integer
    Dim ErrMsg As String

    intFileNum = FreeFile
    On Error GoTo ErrorHandler
    Open pFolder & pFileName For Append As intFileNum
    
    Print #intFileNum, pMessage
    
    Close intFileNum
ErrorHandler:
    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Decode_CST_Information_Elements", ErrMsg)
    Call Show_Message("Write index error", ErrMsg)
End Function
Public Function Check_Network() As Boolean

  Dim TRasCon(255) As RASCONN95
  Dim lg As Long
  Dim lpcon As Long
  Dim RetVal As Long
  Dim Tstatus As RASCONNSTATUS95
  TRasCon(0).dwSize = 412
  lg = 256 * TRasCon(0).dwSize
  RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
  If RetVal <> 0 Then
'    MsgBox "错误"
    Exit Function
  End If
  Tstatus.dwSize = 160
  RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
  If Tstatus.RasConnState = &H2000 Then
    Check_Network = True
  Else
    Check_Network = False
  End If

'   Check_Network = My.Computer.Network.IsAvailable
End Function


