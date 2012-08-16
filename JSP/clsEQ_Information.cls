VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEQ_Information"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim PATTERN_LIST()                      As PATTERN_LIST_DATA
Dim CONTROL_VALUE()                     As ITEM_CONTROL

Dim BACKLIGHT_VALUE                     As String
Dim PG_PANEL_DRIVE_TYPE                 As String
Dim CURRENT_PFCD                        As String
Dim PRE_PFCD                            As String
Dim CURRENT_PROCESSID                   As String
Dim MES_DATA_EXIST_FLAG                 As String
Dim JOB_DATA_EXIST_FLAG                 As String
Dim SHARED_EXIST_FLAG                   As String
Dim MES_DATA_COMMAND                    As String
Dim JOB_DATA_COMMAND                    As String
Dim SHARE_DATA_COMMAND                  As String
Dim QDAC_COMMAND                        As String
Dim RBBC_COMMAND                        As String
Dim RABC_COMMAND                        As String

Dim RE_CONTACT_FLAG                     As Boolean
Dim RE_ALIGNMENT_FLAG                   As Boolean
Dim DEFECT_FILE_UPLOAD                  As Boolean

Dim PATTERN_COUNT                       As Integer
Dim PG_PORTID                           As Integer
Dim CONTROL_COUNT                       As Integer
Dim LOI_STEP                            As Integer

Public Sub Set_BackLight_Value(ByVal pValue As String)

    BACKLIGHT_VALUE = pValue
    
End Sub

Public Function Get_BackLight_Value() As String

    Get_BackLight_Value = BACKLIGHT_VALUE
    
End Function

Public Sub Read_PATTERN_LIST(ByVal pFileName As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim typPATTERN_LIST             As PATTERN_LIST_DATA
    
    Dim strPath                     As String
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    Dim intPattern_Index            As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    strPath = App.PATH & "\Env\Standard_Info\"
    If Dir(strPath & pFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & pFileName For Input As intFileNum
        
        intPattern_Index = 0
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "DELETE * FROM PATTERN_LIST"
            dbMyDB.Execute (strQuery)
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                
                intPos = InStr(strTemp, ",")
                If intPos > 0 Then
                    intPattern_Index = intPattern_Index + 1
                    If UCase(Left(strTemp, intPos - 1)) <> "PATTERN CODE" Then
                        With typPATTERN_LIST
                            .FILENAME = pFileName
                            .PATTERN_CODE = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .PATTERN_NAME = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DELAY_TIME = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .LEVEL = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DH = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .DL = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .VGH = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .VGL = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .RESCUE_HIGH = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            
                            intPos = InStr(strTemp, ",")
                            .RESCUE_LOW = Left(strTemp, intPos - 1)
                            .VCOM = Mid(strTemp, intPos + 1)
                                                    
                            
                            strQuery = "INSERT INTO PATTERN_LIST VALUES ("
                            strQuery = strQuery & "'" & .FILENAME & "', "
                            strQuery = strQuery & "'" & .PATTERN_CODE & "', "
                            strQuery = strQuery & "'" & .PATTERN_NAME & "', "
                            strQuery = strQuery & .DELAY_TIME & ", "
                            strQuery = strQuery & .LEVEL & ", "
                            strQuery = strQuery & .DH & ", "
                            strQuery = strQuery & .DL & ", "
                            strQuery = strQuery & .VGH & ", "
                            strQuery = strQuery & .VGL & ", "
                            strQuery = strQuery & .RESCUE_HIGH & ", "
                            strQuery = strQuery & .RESCUE_LOW & ", "
                            strQuery = strQuery & .VCOM & ")"
                            
                            dbMyDB.Execute (strQuery)
                        End With
                    End If
                End If
            Wend
            
            dbMyDB.Close
        End If
        
        Close intFileNum
        PATTERN_COUNT = intPattern_Index
    Else
        Call Show_Message("File error", pFileName & " does not exist.")
    End If
    
End Sub

Public Sub Set_PATTERN_LIST(ByVal pFileName As String)

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intRecordCount              As Integer
    Dim intRecordIndex              As Integer
    
    intRecordIndex = 0
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PATTERN_LIST WHERE"
        strQuery = strQuery & " FILENAME = '" & pFileName & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            intRecordCount = lstRecord.RecordCount
            ReDim PATTERN_LIST(intRecordCount)
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRecordIndex = intRecordIndex + 1
                With PATTERN_LIST(intRecordIndex)
                    .PATTERN_CODE = lstRecord.Fields("PATTERN_CODE")
                    .PATTERN_NAME = lstRecord.Fields("PATTERN_NAME")
                    .DELAY_TIME = lstRecord.Fields("DELAY_TIME")
                    .LEVEL = lstRecord.Fields("LEVEL")
                    .DH = lstRecord.Fields("DH")
                    .DL = lstRecord.Fields("DL")
                    .VGH = lstRecord.Fields("VGH")
                    .VGL = lstRecord.Fields("VGL")
                    .RESCUE_HIGH = lstRecord.Fields("RESCUE_HIGH")
                    .RESCUE_LOW = lstRecord.Fields("RESCUE_LOW")
                    .VCOM = lstRecord.Fields("VCOM")
                End With
                lstRecord.MoveNext
            Wend
        End If
        PATTERN_COUNT = intRecordCount
        
        lstRecord.Close
        
        dbMyDB.Close
    End If

End Sub

Public Function Get_PATTERN_COUNT() As Integer

    Get_PATTERN_COUNT = PATTERN_COUNT
    
End Function

Public Sub Get_PATTERN_LIST_by_PATTERN_CODE(ByVal pPTN_CODE As String, pPTN_NAME As String, pDELAY_TIME As Integer, pLEVEL As Integer, pDH As Integer, _
                                            pDL As Integer, pVGH As Integer, pVGL As Integer, pRESCUE_HIGH As Integer, pRESCUE_LOW As Integer, pVCOM As Integer)

    Dim intIndex                        As Integer
    
    For intIndex = 1 To PATTERN_COUNT
        If PATTERN_LIST(intIndex).PATTERN_CODE = pPTN_CODE Then
            With PATTERN_LIST(intIndex)
                pPTN_NAME = .PATTERN_NAME
                pDELAY_TIME = .DELAY_TIME
                pLEVEL = .LEVEL
                pDH = .DH
                pDL = .DL
                pVGH = .VGH
                pVGL = .VGL
                pRESCUE_HIGH = .RESCUE_HIGH
                pRESCUE_LOW = .RESCUE_LOW
                pVCOM = .VCOM
            End With
        End If
    Next intIndex
    
End Sub

Public Sub Get_PATTERN_LIST_by_Index(ByVal pIndex As Integer, pPTN_CODE As String, pPTN_NAME As String, pDELAY_TIME As Integer, pLEVEL As Integer, pDH As Integer, _
                                     pDL As Integer, pVGH As Integer, pVGL As Integer, pRESCUE_HIGH As Integer, pRESCUE_LOW As Integer, pVCOM As Integer)

    If PATTERN_COUNT > 0 Then
        With PATTERN_LIST(pIndex)
            pPTN_CODE = .PATTERN_CODE
            pPTN_NAME = .PATTERN_NAME
            pDELAY_TIME = .DELAY_TIME
            pLEVEL = .LEVEL
            pDH = .DH
            pDL = .DL
            pVGH = .VGH
            pVGL = .VGL
            pRESCUE_HIGH = .RESCUE_HIGH
            pRESCUE_LOW = .RESCUE_LOW
            pVCOM = .VCOM
        End With
    Else
        pPTN_CODE = ""
        pPTN_NAME = ""
        pDELAY_TIME = 0
        pLEVEL = 0
        pDH = 0
        pDL = 0
        pVGH = 0
        pVGL = 0
        pRESCUE_HIGH = 0
        pRESCUE_LOW = 0
        pVCOM = 0
    End If
    
End Sub

Public Sub Set_PATTERN_START_by_Index(ByVal pIndex As Integer)

    If PATTERN_COUNT > 0 Then
        PATTERN_LIST(pIndex).INSPECTION_START = Timer
    End If
    
End Sub

Public Sub Set_PATTERN_END_by_Index(ByVal pIndex As Integer)

    If PATTERN_COUNT > 0 Then
        PATTERN_LIST(pIndex).INSPECTION_END = Timer
        
        PATTERN_LIST(pIndex).INSPECTION_TIME = PATTERN_LIST(pIndex).INSPECTION_END - PATTERN_LIST(pIndex).INSPECTION_START
    End If
    
End Sub

Public Function Get_PATTERN_INSPECTION_TIME_by_Index(ByVal pIndex As Integer) As Double

    If PATTERN_COUNT > 0 Then
        Get_PATTERN_INSPECTION_TIME_by_Index = PATTERN_LIST(pIndex).INSPECTION_TIME
    End If
    
End Function

Public Sub Set_PG_Panel_Drive_Type(ByVal pDriveType As String)

    PG_PANEL_DRIVE_TYPE = pDriveType
    
End Sub

Public Function Get_PG_Panel_Drive_Type() As String

    Get_PG_Panel_Drive_Type = PG_PANEL_DRIVE_TYPE
    
End Function

Public Sub Set_PG_PortID(ByVal pPortID As Integer)

    PG_PORTID = pPortID
    
End Sub

Public Function Get_PG_PortID() As Integer

    Get_PG_PortID = PG_PORTID
    
End Function

Public Sub Set_Current_PFCD(ByVal pPFCD As String)

    PRE_PFCD = CURRENT_PFCD
    CURRENT_PFCD = pPFCD
    
End Sub

Public Function Get_Current_PFCD() As String

    Get_Current_PFCD = CURRENT_PFCD
    
End Function

Public Function Get_Pre_PFCD() As String

    Get_Pre_PFCD = PRE_PFCD
    
End Function

Public Sub Set_Control_Data()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intIndex                    As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM ITEM_CONTROL"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            CONTROL_COUNT = lstRecord.RecordCount
            ReDim CONTROL_VALUE(CONTROL_COUNT)
            lstRecord.MoveFirst
            intIndex = 0
            While lstRecord.EOF = False
                intIndex = intIndex + 1
                With CONTROL_VALUE(intIndex)
                    .ITEM_NAME = lstRecord.Fields("ITEM_NAME")
                    .ENABLE_DISABLE = lstRecord.Fields("USES")
                End With
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("Set_Control_Data", ErrMsg)
    
End Sub

Public Function Get_Control_Data(ByVal pITEM_NAME As String) As String

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intIndex                    As Integer
    
    Dim ErrMsg                      As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "STANDARD_INFO.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM ITEM_CONTROL WHERE "
        strQuery = strQuery & "ITEM_NAME = '" & pITEM_NAME & "'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            Get_Control_Data = lstRecord.Fields("USES")
        Else
            Get_Control_Data = "D"
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
End Function

Public Sub Set_MES_Data_for_API(ByVal pMES_DATA_EXIST As String, ByVal pJOB_DATA_EXIST As String, ByVal pSHARE_DATA_EXIST As String, _
                                ByVal pMES_DATA As String, ByVal pJOB_DATA As String, ByVal pSHARE_DATA As String)
                                
    MES_DATA_EXIST_FLAG = pMES_DATA_EXIST
    JOB_DATA_EXIST_FLAG = pJOB_DATA_EXIST
    SHARED_EXIST_FLAG = pSHARE_DATA_EXIST
    
    MES_DATA_COMMAND = pMES_DATA
    JOB_DATA_COMMAND = pJOB_DATA
    SHARE_DATA_COMMAND = pSHARE_DATA
    
End Sub

Public Sub Get_MES_Data_for_API(pMES_DATA_EXIST As String, pJOB_DATA_EXIST As String, pSHARE_DATA_EXIST As String, _
                                pMES_DATA As String, pJOB_DATA As String, pSHARE_DATA As String)

    pMES_DATA_EXIST = MES_DATA_EXIST_FLAG
    pJOB_DATA_EXIST = JOB_DATA_EXIST_FLAG
    pSHARE_DATA_EXIST = SHARED_EXIST_FLAG
    
    pMES_DATA = MES_DATA_COMMAND
    pJOB_DATA = JOB_DATA_COMMAND
    pSHARE_DATA = SHARE_DATA_COMMAND
    
End Sub

Public Sub Set_LOI_STEP(ByVal pSTEP As Integer)

    LOI_STEP = pSTEP
    
End Sub

Public Function Get_LOI_STEP() As Integer

    Get_LOI_STEP = LOI_STEP
    
End Function

Public Sub Set_QDAC_COMMAND(ByVal pCommand As String)

    QDAC_COMMAND = pCommand
    
End Sub

Public Function Get_QDAC_COMMAND() As String

    Get_QDAC_COMMAND = QDAC_COMMAND
    
End Function

Public Sub Set_Re_Contact_Flag(ByVal pFlag As Boolean)

    RE_CONTACT_FLAG = pFlag
    
End Sub

Public Function Get_Re_Contact_Flag() As Boolean

    Get_Re_Contact_Flag = False
    
End Function

Public Sub set_Re_Alignment_Flag(ByVal pFlag As Boolean)

    RE_ALIGNMENT_FLAG = pFlag
    
End Sub

Public Function Get_Re_Alignment_Flag() As Boolean

    Get_Re_Alignment_Flag = RE_ALIGNMENT_FLAG
    
End Function

Public Sub Set_DEFECT_UPLOAD(ByVal pDEFECT_UPLOAD As Boolean)

    DEFECT_FILE_UPLOAD = pDEFECT_UPLOAD
    
End Sub

Public Function Get_DEFECT_UPLOAD() As Boolean

    Get_DEFECT_UPLOAD = DEFECT_FILE_UPLOAD
    
End Function

Public Sub Set_RBBC_Command(ByVal pCommand As String)

    RBBC_COMMAND = pCommand
    
End Sub

Public Function Get_RBBC_Command() As String

    Get_RBBC_Command = RBBC_COMMAND

End Function

Public Sub Set_RABC_Command(ByVal pCommand As String)

    RABC_COMMAND = pCommand
    
End Sub

Public Function Get_RABC_Command() As String

    Get_RABC_Command = RABC_COMMAND
    
End Function

'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - Add class function for save current process number
'
'==========================================================================================================
Public Sub Set_Current_PROCESSID(ByVal pPROCESSID As String)

    CURRENT_PROCESSID = pPROCESSID
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2011. 12. 26
'  Modify by K.H. KIM
'  Content
'    - Add class function for read current process number
'
'==========================================================================================================
Public Function Get_Current_PROCESSID() As String

    Get_Current_PROCESSID = CURRENT_PROCESSID
    
End Function
